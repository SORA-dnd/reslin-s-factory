/**
 * Reslin's Factory - AI Automation (Global Edition)
 * 
 * Features:
 * - Patreon OAuth Authentication (Moulinette-style)
 * - User-provided API Key (OpenAI-compatible)
 * - 5-Step RAG Architecture via Cloudflare Worker
 * - Encrypted membership verification
 */

const MODULE_ID = 'reslin-s-factory';
const WORKER_URL = 'https://soramod.online'; // Same worker, different auth mode

// Patreon OAuth Configuration
const PATREON_CONFIG = {
    clientId: 'egcC6ZbmHZ41yC6AKqQc2RJZJomKu9bYS8bUKnYReAL8F3nC8YLa-fzNISeBoQhE',
    redirectUri: 'https://soramod.online/patreon/callback',
    scope: 'identity identity.memberships',
    authUrl: 'https://www.patreon.com/oauth2/authorize'
};

// ================= INITIALIZATION =================

Hooks.once('init', () => {
    console.log(`${MODULE_ID} | Initializing Reslin's Factory - AI Automation (Global Edition)`);

    // Patreon Token (encrypted, stored locally)
    game.settings.register(MODULE_ID, 'patreonToken', {
        name: game.i18n.localize('RESLIN.Settings.PatreonToken.Name'),
        hint: game.i18n.localize('RESLIN.Settings.PatreonToken.Hint'),
        scope: 'client',
        config: false, // Hidden, managed via UI
        type: String,
        default: '',
    });

    // Patreon Session ID (for verification)
    game.settings.register(MODULE_ID, 'patreonSessionId', {
        scope: 'client',
        config: false,
        type: String,
        default: '',
    });

    // Membership Status Cache
    game.settings.register(MODULE_ID, 'membershipStatus', {
        scope: 'client',
        config: false,
        type: Object,
        default: { valid: false, tier: null, expiresAt: null, lastCheck: null },
    });

    // User API Key
    game.settings.register(MODULE_ID, 'apiKey', {
        name: game.i18n.localize('RESLIN.Settings.ApiKey.Name'),
        hint: game.i18n.localize('RESLIN.Settings.ApiKey.Hint'),
        scope: 'client',
        config: true,
        type: String,
        default: '',
    });

    // User API Base URL
    game.settings.register(MODULE_ID, 'apiBaseUrl', {
        name: game.i18n.localize('RESLIN.Settings.ApiBaseUrl.Name'),
        hint: game.i18n.localize('RESLIN.Settings.ApiBaseUrl.Hint'),
        scope: 'client',
        config: true,
        type: String,
        default: 'https://api.openai.com/v1',
    });

    // User API Model
    game.settings.register(MODULE_ID, 'apiModel', {
        name: game.i18n.localize('RESLIN.Settings.ApiModel.Name'),
        hint: game.i18n.localize('RESLIN.Settings.ApiModel.Hint'),
        scope: 'client',
        config: true,
        type: String,
        default: 'gpt-4',
    });
});

Hooks.on('renderItemDirectory', (app, html, data) => {
    if (!game.user.isGM) return;

    const $html = html instanceof HTMLElement ? $(html) : html;
    const button = $(`<button class="create-ai-item-global"><i class="fab fa-patreon"></i> AI Generate (Global)</button>`);
    button.on('click', () => {
        new ReslinFactoryDialog().render(true);
    });

    $html.find('.directory-header .header-actions').append(button);
});

// ================= PATREON AUTHENTICATION =================

class PatreonAuth {
    static generateState() {
        return foundry.utils.randomID(32);
    }

    static async initiateOAuth() {
        const state = this.generateState();
        
        // Store state for verification
        sessionStorage.setItem('reslin_patreon_state', state);
        
        const params = new URLSearchParams({
            response_type: 'code',
            client_id: PATREON_CONFIG.clientId,
            redirect_uri: PATREON_CONFIG.redirectUri,
            scope: PATREON_CONFIG.scope,
            state: state
        });

        const authUrl = `${PATREON_CONFIG.authUrl}?${params.toString()}`;
        
        // Try to open popup for OAuth
        let popup = null;
        try {
            popup = window.open(authUrl, 'patreon_auth', 'width=600,height=700');
        } catch (e) {
            console.warn('Popup open failed:', e);
        }
        
        // Check if popup was blocked
        if (!popup || popup.closed || typeof popup.closed === 'undefined') {
            // Fallback: Open in new tab and ask user to paste code
            const useNewTab = await Dialog.confirm({
                title: "Popup Blocked",
                content: `<p>Your browser blocked the popup window.</p>
                         <p>Would you like to open Patreon authorization in a new tab instead?</p>
                         <p><small>After authorizing, you'll need to copy the authorization code from the URL.</small></p>`,
                yes: () => true,
                no: () => false
            });
            
            if (useNewTab) {
                window.open(authUrl, '_blank');
                
                // Ask user to paste the code
                const code = await this.promptForCode();
                if (code) {
                    return await this.exchangeCode(code);
                }
            }
            
            throw new Error("Authentication cancelled - popup blocked");
        }
        
        // Listen for callback
        return new Promise((resolve, reject) => {
            let resolved = false;
            
            const checkClosed = setInterval(() => {
                try {
                    if (!popup || popup.closed) {
                        clearInterval(checkClosed);
                        if (!resolved) {
                            // Check if we got a token
                            const token = game.settings.get(MODULE_ID, 'patreonToken');
                            if (token) {
                                resolved = true;
                                resolve(token);
                            } else {
                                reject(new Error('Authentication cancelled'));
                            }
                        }
                    }
                } catch (e) {
                    // Cross-origin error, popup is still open
                }
            }, 500);

            // Also listen for postMessage from callback
            const messageHandler = async (event) => {
                // Accept messages from our worker domain
                if (!event.origin.includes('soramod.online')) return;
                
                if (event.data.type === 'patreon-oauth-callback' && event.data.code) {
                    clearInterval(checkClosed);
                    window.removeEventListener('message', messageHandler);
                    
                    try {
                        if (popup && !popup.closed) popup.close();
                    } catch (e) {}
                    
                    if (!resolved) {
                        resolved = true;
                        try {
                            const token = await this.exchangeCode(event.data.code);
                            resolve(token);
                        } catch (e) {
                            reject(e);
                        }
                    }
                }
            };
            
            window.addEventListener('message', messageHandler);
            
            // Timeout after 5 minutes
            setTimeout(() => {
                if (!resolved) {
                    clearInterval(checkClosed);
                    window.removeEventListener('message', messageHandler);
                    reject(new Error('Authentication timed out'));
                }
            }, 300000);
        });
    }

    static async promptForCode() {
        return new Promise((resolve) => {
            new Dialog({
                title: "Enter Authorization Code",
                content: `
                    <p>After authorizing on Patreon, look at the URL in your browser.</p>
                    <p>It should look like: <code>https://soramod.online/patreon/callback?code=XXXXX&state=...</code></p>
                    <p>Copy the <strong>code</strong> value and paste it below:</p>
                    <input type="text" id="patreon-auth-code" style="width: 100%; margin-top: 10px;" placeholder="Paste authorization code here">
                `,
                buttons: {
                    submit: {
                        icon: '<i class="fas fa-check"></i>',
                        label: "Submit",
                        callback: (html) => {
                            const code = html.find('#patreon-auth-code').val();
                            resolve(code || null);
                        }
                    },
                    cancel: {
                        icon: '<i class="fas fa-times"></i>',
                        label: "Cancel",
                        callback: () => resolve(null)
                    }
                },
                default: "submit"
            }).render(true);
        });
    }

    static async exchangeCode(code) {
        // Exchange code for token via worker (keeps client_secret secure)
        const response = await fetch(`${WORKER_URL}/patreon/exchange`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ code, redirect_uri: PATREON_CONFIG.redirectUri })
        });

        if (!response.ok) {
            throw new Error('Failed to exchange authorization code');
        }

        const data = await response.json();
        
        // Store encrypted token
        await game.settings.set(MODULE_ID, 'patreonToken', data.access_token);
        await game.settings.set(MODULE_ID, 'patreonSessionId', data.session_id || foundry.utils.randomID());
        
        return data.access_token;
    }

    static async verifyMembership() {
        const token = game.settings.get(MODULE_ID, 'patreonToken');
        if (!token) return { valid: false, reason: 'No token' };

        const sessionId = game.settings.get(MODULE_ID, 'patreonSessionId');
        
        try {
            const response = await fetch(`${WORKER_URL}/patreon/verify`, {
                method: 'POST',
                headers: { 
                    'Content-Type': 'application/json',
                    'x-patreon-token': token,
                    'x-session-id': sessionId
                }
            });

            if (!response.ok) {
                const err = await response.json();
                return { valid: false, reason: err.error || 'Verification failed' };
            }

            const data = await response.json();
            
            // Cache membership status (including quota for free tier)
            const status = {
                valid: data.valid,
                tier: data.tier,
                expiresAt: data.expiresAt,
                lastCheck: Date.now(),
                quota: data.quota || null,  // Save quota info for free tier users
                whitelisted: data.whitelisted || false
            };
            await game.settings.set(MODULE_ID, 'membershipStatus', status);
            
            return status;
        } catch (e) {
            console.error(`${MODULE_ID} | Membership verification failed:`, e);
            return { valid: false, reason: e.message };
        }
    }

    static async disconnect() {
        await game.settings.set(MODULE_ID, 'patreonToken', '');
        await game.settings.set(MODULE_ID, 'patreonSessionId', '');
        await game.settings.set(MODULE_ID, 'membershipStatus', { valid: false, tier: null, expiresAt: null, lastCheck: null });
    }

    static isConnected() {
        return !!game.settings.get(MODULE_ID, 'patreonToken');
    }

    static getCachedStatus() {
        const status = game.settings.get(MODULE_ID, 'membershipStatus');
        // Cache valid for 1 hour
        if (status.lastCheck && (Date.now() - status.lastCheck) < 3600000) {
            return status;
        }
        return null;
    }

    static getQuotaInfo() {
        const status = game.settings.get(MODULE_ID, 'membershipStatus');
        return status.quota || null;
    }
}

// ================= MAIN DIALOG =================

class ReslinFactoryDialog extends Application {
    constructor(options) {
        super(options);
        this._fixContext = {
            item: null,
            journal: null
        };
    }

    static get defaultOptions() {
        return mergeObject(super.defaultOptions, {
            id: 'reslin-factory-dialog',
            title: game.i18n.localize('RESLIN.Dialog.Title'),
            template: `modules/${MODULE_ID}/templates/generator-dialog.html`,
            width: 900,
            height: 750,
            classes: ['reslin-factory-dialog'],
            resizable: true,
            dragDrop: [{ dragSelector: null, dropSelector: ".drop-zone" }]
        });
    }

    getData() {
        return {
            isConnected: PatreonAuth.isConnected(),
            apiKey: game.settings.get(MODULE_ID, 'apiKey'),
            apiBaseUrl: game.settings.get(MODULE_ID, 'apiBaseUrl'),
            apiModel: game.settings.get(MODULE_ID, 'apiModel')
        };
    }

    activateListeners(html) {
        super.activateListeners(html);
        
        // Navigation
        html.find('.nav-btn').on('click', this._onNavClick.bind(this));
        
        // Mode Selection
        html.find('.mode-card').on('click', this._onModeClick.bind(this));
        
        // Patreon Connect
        html.find('#patreon-connect-btn').on('click', this._onPatreonConnect.bind(this));
        
        // API Config Changes
        html.find('#api-base-url').on('change', (e) => game.settings.set(MODULE_ID, 'apiBaseUrl', e.target.value));
        html.find('#api-key').on('change', (e) => game.settings.set(MODULE_ID, 'apiKey', e.target.value));
        html.find('#api-model').on('change', (e) => game.settings.set(MODULE_ID, 'apiModel', e.target.value));
        
        // Generate
        html.find('#generate-btn').on('click', this._onGenerate.bind(this));
        
        // Fix
        html.find('#fix-btn').on('click', this._onFix.bind(this));
        
        // Drop Zone Remove
        html.find('.remove-btn').on('click', this._onRemoveDrop.bind(this));

        // Initialize UI state
        this._initializeUI(html);
    }

    async _initializeUI(html) {
        // Load saved API config
        html.find('#api-base-url').val(game.settings.get(MODULE_ID, 'apiBaseUrl'));
        html.find('#api-key').val(game.settings.get(MODULE_ID, 'apiKey'));
        html.find('#api-model').val(game.settings.get(MODULE_ID, 'apiModel'));

        // Update Patreon status
        await this._updatePatreonStatus(html);
    }

    async _updatePatreonStatus(html) {
        const badge = html.find('#patreon-status-badge');
        const btn = html.find('#patreon-connect-btn');
        const settingsStatus = html.find('#settings-patreon-status');
        const membershipStatus = html.find('#settings-membership-status');
        const membershipExpires = html.find('#settings-membership-expires');
        const quotaDisplay = html.find('#quota-display');
        const quotaRemaining = html.find('#quota-remaining');
        const quotaLimit = html.find('#quota-limit');

        // Hide quota display by default
        quotaDisplay.addClass('hidden').removeClass('warning exhausted');

        if (!PatreonAuth.isConnected()) {
            badge.removeClass('connected verifying').addClass('disconnected').text('Not Connected');
            btn.removeClass('disconnect').html('<i class="fab fa-patreon"></i> Connect Patreon');
            settingsStatus.text('Not Connected');
            membershipStatus.text('-');
            membershipExpires.text('-');
            return;
        }

        // Check cached status first
        let status = PatreonAuth.getCachedStatus();
        
        if (!status) {
            badge.removeClass('connected disconnected').addClass('verifying').text('Verifying...');
            status = await PatreonAuth.verifyMembership();
        }

        if (status.valid) {
            badge.removeClass('disconnected verifying').addClass('connected').text('Active Member');
            btn.addClass('disconnect').html('<i class="fas fa-unlink"></i> Disconnect');
            settingsStatus.text('Connected');
            
            // Handle different tiers
            if (status.whitelisted) {
                membershipStatus.text('Whitelist (Unlimited)');
                membershipExpires.text('∞');
            } else if (status.tier === 'free' && status.quota) {
                // Free tier - show quota
                membershipStatus.text('Free Tier');
                membershipExpires.text('Monthly Reset');
                
                // Show quota display
                const remaining = status.quota.remaining;
                const limit = status.quota.limit;
                quotaRemaining.text(remaining);
                quotaLimit.text(limit);
                quotaDisplay.removeClass('hidden');
                
                // Add warning/exhausted classes based on remaining
                if (remaining === 0) {
                    quotaDisplay.addClass('exhausted');
                } else if (remaining <= 2) {
                    quotaDisplay.addClass('warning');
                }
            } else {
                // Paid tier
                membershipStatus.text(status.tier || 'Paid');
                membershipExpires.text(status.expiresAt ? new Date(status.expiresAt).toLocaleDateString() : 'Active');
            }
        } else {
            badge.removeClass('connected verifying').addClass('disconnected').text('Membership Required');
            btn.removeClass('disconnect').html('<i class="fab fa-patreon"></i> Reconnect');
            settingsStatus.text('Connected (No Active Membership)');
            membershipStatus.text('Inactive');
            membershipExpires.text('-');
        }
    }

    _onNavClick(event) {
        event.preventDefault();
        const btn = $(event.currentTarget);
        const view = btn.data('view');
        
        this.element.find('.nav-btn').removeClass('active');
        btn.addClass('active');
        
        this.element.find('.view-panel').addClass('hidden');
        this.element.find(`#view-${view}`).removeClass('hidden');
    }

    _onModeClick(event) {
        event.preventDefault();
        const card = $(event.currentTarget);
        const value = card.data('value');
        
        this.element.find('.mode-card').removeClass('active');
        card.addClass('active');
        this.element.find('#generation-mode').val(value);
    }

    async _onPatreonConnect(event) {
        event.preventDefault();
        const btn = $(event.currentTarget);

        if (PatreonAuth.isConnected() && btn.hasClass('disconnect')) {
            await PatreonAuth.disconnect();
            await this._updatePatreonStatus(this.element);
            ui.notifications.info('Patreon account disconnected.');
            return;
        }

        try {
            btn.prop('disabled', true).html('<i class="fas fa-spinner fa-spin"></i> Connecting...');
            await PatreonAuth.initiateOAuth();
            await this._updatePatreonStatus(this.element);
            ui.notifications.info('Patreon account connected successfully!');
        } catch (e) {
            console.error(e);
            ui.notifications.error(`Patreon connection failed: ${e.message}`);
        } finally {
            btn.prop('disabled', false);
            await this._updatePatreonStatus(this.element);
        }
    }

    async _onDrop(event) {
        event.preventDefault();
        const data = TextEditor.getDragEventData(event);
        const dropZone = $(event.currentTarget);
        const type = dropZone.data('type');

        if (data.type !== type) {
            ui.notifications.warn(`Please drop a ${type} here.`);
            return;
        }

        let document;
        try {
            if (type === 'Item') {
                document = await Item.fromDropData(data);
            } else if (type === 'JournalEntry') {
                document = await JournalEntry.fromDropData(data);
            }
        } catch (e) {
            console.error(e);
            return;
        }

        if (!document) return;

        if (type === 'Item') this._fixContext.item = document;
        if (type === 'JournalEntry') this._fixContext.journal = document;

        const content = dropZone.find('.drop-content');
        const preview = dropZone.find('.drop-preview');
        
        content.addClass('hidden');
        preview.removeClass('hidden');
        
        preview.find('.preview-name').text(document.name);
        if (type === 'Item') {
            preview.find('.preview-img').attr('src', document.img);
        }
    }

    _onRemoveDrop(event) {
        event.stopPropagation();
        const btn = $(event.currentTarget);
        const dropZone = btn.closest('.drop-zone');
        const type = dropZone.data('type');

        if (type === 'Item') this._fixContext.item = null;
        if (type === 'JournalEntry') this._fixContext.journal = null;

        dropZone.find('.drop-content').removeClass('hidden');
        dropZone.find('.drop-preview').addClass('hidden');
    }

    async _onFix(event) {
        event.preventDefault();
        ui.notifications.info("Interactive Fix functionality coming soon!");
    }

    _log(message, type = 'info') {
        const consoleOutput = this.element.find('#console-output');
        const timestamp = new Date().toLocaleTimeString();
        let colorClass = 'log-info';
        if (type === 'warn') colorClass = 'log-warn';
        if (type === 'error') colorClass = 'log-error';
        if (type === 'success') colorClass = 'log-success';

        consoleOutput.append(`<div class="log-entry ${colorClass}">[${timestamp}] ${message}</div>`);
        consoleOutput.scrollTop(consoleOutput[0].scrollHeight);
    }

    _updateStatus(text) {
        this.element.find('#status-display .status-text').text(text);
    }

    _showPhaseSummary(phaseName, summary, progress = 0) {
        const container = this.element.find('#phase-summary-container');
        const title = this.element.find('#phase-summary-phase');
        const content = this.element.find('#phase-summary-text');
        const progressBar = this.element.find('#phase-progress-fill');
        const progressText = this.element.find('#phase-progress-text');

        if (container.hasClass('hidden')) container.removeClass('hidden');

        title.text(phaseName);
        content.html(summary);
        progressBar.css('width', `${progress}%`);
        progressText.text(`${progress}%`);
    }

    async _callWorker(endpoint, payload) {
        // Get Patreon token for authentication
        const patreonToken = game.settings.get(MODULE_ID, 'patreonToken');
        if (!patreonToken) {
            throw new Error(game.i18n.localize('RESLIN.Error.NoToken'));
        }

        // Get user API configuration
        const userApiKey = game.settings.get(MODULE_ID, 'apiKey');
        const userApiBase = game.settings.get(MODULE_ID, 'apiBaseUrl');
        const userApiModel = game.settings.get(MODULE_ID, 'apiModel');

        if (!userApiKey) {
            throw new Error(game.i18n.localize('RESLIN.Error.NoApiKey'));
        }
        if (!userApiBase) {
            throw new Error(game.i18n.localize('RESLIN.Error.NoApiUrl'));
        }

        const sessionId = game.settings.get(MODULE_ID, 'patreonSessionId');

        const response = await fetch(`${WORKER_URL}/${endpoint}`, {
            method: 'POST',
            headers: { 
                'Content-Type': 'application/json',
                'x-patreon-token': patreonToken,
                'x-session-id': sessionId,
                'x-user-api-key': userApiKey,
                'x-user-api-base': userApiBase,
                'x-user-api-model': userApiModel
            },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            const errText = await response.text();
            let errMsg = errText;
            try {
                const errJson = JSON.parse(errText);
                if (errJson.error) errMsg = errJson.error;
            } catch(e) {}
            throw new Error(`Worker Error (${endpoint}): ${errMsg}`);
        }

        return await response.json();
    }

    async _onGenerate(event) {
        event.preventDefault();
        const description = this.element.find('#item-description').val();
        const mode = this.element.find('#generation-mode').val();

        // Reset UI
        this.element.find('#console-output').empty();
        this.element.find('#phase-summary-container').addClass('hidden');

        if (!description) {
            ui.notifications.warn("Please enter an item description.");
            return;
        }

        // Verify membership first
        this._updateStatus("VERIFYING MEMBERSHIP...");
        this._log("Verifying Patreon membership...", 'info');

        const status = await PatreonAuth.verifyMembership();
        if (!status.valid) {
            this._updateStatus("MEMBERSHIP REQUIRED");
            this._log(`Membership verification failed: ${status.reason}`, 'error');
            ui.notifications.error(game.i18n.localize('RESLIN.Error.MembershipRequired'));
            return;
        }

        this._log("✓ Membership verified!", 'success');
        this._updateStatus("INITIALIZING...");
        this._log("Starting 5-Phase Cloud Generation (Global)...", 'info');

        try {
            // ===== PHASE 1: FEATURE EXTRACTION =====
            this._log("PHASE 1: Analyzing request...", 'info');
            this._updateStatus("PHASE 1: ANALYSIS");
            
            const step1Res = await this._callWorker('step1_extract', { description });
            const features = step1Res;
            
            this._showPhaseSummary(
                "Phase 1: Analysis",
                `<strong>Intent:</strong> ${features.intent}<br><strong>Complexity:</strong> ${features.complexity}<br><strong>Tags:</strong> ${features.tags.join(', ')}`,
                20
            );
            this._log(`✓ Features extracted: ${features.tags.join(', ')}`, 'success');

            // ===== PHASE 2: MACRO RETRIEVAL =====
            this._log("PHASE 2: Retrieving Reference Macros...", 'info');
            this._updateStatus("PHASE 2: MACROS");

            const step2Res = await this._callWorker('step2_macros', { features });
            const referenceCode = step2Res.referenceCode || "";

            this._showPhaseSummary(
                "Phase 2: Macro Retrieval",
                `<strong>Reference Code:</strong> ${referenceCode ? 'Found' : 'None'}<br><strong>Source:</strong> Cloud KV`,
                40
            );

            // ===== PHASE 3: KNOWLEDGE RETRIEVAL =====
            this._log("PHASE 3: Retrieving Knowledge Base...", 'info');
            this._updateStatus("PHASE 3: KNOWLEDGE");

            const step3Res = await this._callWorker('step3_knowledge', { features, description });
            const knowledge = step3Res.knowledge || "";

            this._showPhaseSummary(
                "Phase 3: Knowledge Retrieval",
                `<strong>Knowledge Docs:</strong> ${knowledge.length > 100 ? 'Loaded' : 'Empty'}<br><strong>RAG Status:</strong> Active`,
                60
            );

            // ===== PHASE 4: GENERATION =====
            this._log("PHASE 4: Generating Content (Using Your API)...", 'info');
            this._updateStatus("PHASE 4: GENERATION");

            const step4Res = await this._callWorker('step4_generate', { 
                description, 
                knowledge, 
                features, 
                referenceCode, 
                mode 
            });
            
            let itemJson, macroCode, tutorial;

            if (step4Res.result) {
                const parsed = JSON.parse(step4Res.result);
                itemJson = parsed.itemJson;
                macroCode = parsed.macroCode;
                tutorial = parsed.tutorial;
            } else {
                itemJson = step4Res.itemJson;
                macroCode = step4Res.macroCode;
                tutorial = step4Res.tutorial;
            }

            this._showPhaseSummary(
                "Phase 4: Generation",
                `<strong>Item:</strong> ${itemJson.name}<br><strong>Macro:</strong> ${macroCode ? 'Generated' : 'N/A'}`,
                80
            );
            this._log(`✓ Content generated: "${itemJson.name}"`, 'success');

            // ===== PHASE 5: FINAL FIX =====
            this._log("PHASE 5: Final Fix & Validation...", 'info');
            this._updateStatus("PHASE 5: FINAL FIX");

            const step5Res = await this._callWorker('step5_fix', { itemJson, macroCode, tutorial });
            if (step5Res.fixedJson) {
                itemJson = step5Res.fixedJson;
                this._log("✓ AI applied final fixes.", 'success');
            }

            // Local Validation
            itemJson = this._cleanIds(itemJson);
            itemJson = this._validateAndFixItem(itemJson);

            this._showPhaseSummary(
                "Phase 5: Completion",
                `<strong>Status:</strong> Ready<br><strong>Pipeline:</strong> 5-Step Cloud Execution`,
                100
            );

            // ===== COMPLETION =====
            if (macroCode || tutorial) {
                await this._createItemDocumentation(itemJson, macroCode, tutorial);
            }

            const item = await Item.create(itemJson);
            item.sheet.render(true);

            this._updateStatus("COMPLETE");
            this._log(`✓ Item created successfully!`, 'success');
            ui.notifications.info(`Item "${item.name}" created!`);

            setTimeout(() => {
                this.element.find('#phase-summary-container').addClass('hidden');
            }, 5000);

        } catch (error) {
            console.error(error);
            this._updateStatus("ERROR");
            this._log(`Error: ${error.message}`, 'error');
            ui.notifications.error(`Generation failed: ${error.message}`);
        }
    }

    _cleanIds(obj) {
        if (Array.isArray(obj)) {
            return obj.map(v => this._cleanIds(v));
        } else if (obj !== null && typeof obj === 'object') {
            return Object.keys(obj).reduce((acc, key) => {
                if (key !== '_id') {
                    acc[key] = this._cleanIds(obj[key]);
                }
                return acc;
            }, {});
        }
        return obj;
    }

    _validateAndFixItem(itemJson) {
        if (itemJson.system?.activities) {
            const cleanActivities = {};
            for (const [key, activity] of Object.entries(itemJson.system.activities)) {
                if (!activity) continue;
                const newId = foundry.utils.randomID();
                const cleanActivity = duplicate(activity);
                cleanActivity._id = newId;
                
                if (cleanActivity.damage?.parts) {
                    cleanActivity.damage.parts.forEach(part => {
                        if (part.scaling?.mode === 'level') part.scaling.mode = 'whole';
                    });
                }
                
                cleanActivities[newId] = cleanActivity;
            }
            itemJson.system.activities = cleanActivities;
        }
        return itemJson;
    }

    async _createItemDocumentation(item, code, tutorial) {
        try {
            const folderName = "AI Generated Macros (Global)";
            let folder = game.folders.find(f => f.name === folderName && f.type === "JournalEntry");
            if (!folder) {
                folder = await Folder.create({ name: folderName, type: "JournalEntry" });
            }

            const pages = [];
            if (code) {
                pages.push({
                    name: "Macro Code",
                    type: "text",
                    text: { content: `<pre><code>${code}</code></pre>`, format: 1 }
                });
            }
            if (tutorial) {
                pages.push({
                    name: "AI Tutorial",
                    type: "text",
                    text: { content: tutorial, format: 1 }
                });
            }

            await JournalEntry.create({
                name: item.name + " (Documentation)",
                folder: folder.id,
                pages: pages
            });
        } catch (e) {
            console.error("Failed to create documentation:", e);
        }
    }
}
