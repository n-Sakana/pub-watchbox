// bridge.js - Shared C# <-> JS message protocol
const Bridge = {
    _handlers: {},
    _pending: {},
    _nextId: 1,

    // Register a handler for an action from C#
    on(action, callback) {
        this._handlers[action] = callback;
    },

    // Send a message to C# (fire-and-forget)
    send(action, data) {
        if (!window.chrome || !window.chrome.webview) return;
        window.chrome.webview.postMessage({ action, data: data || {} });
    },

    // Send a message to C# and await response
    request(action, data) {
        return new Promise((resolve) => {
            const id = this._nextId++;
            this._pending[id] = resolve;
            if (!window.chrome || !window.chrome.webview) {
                resolve(null);
                return;
            }
            window.chrome.webview.postMessage({
                action, data: data || {}, _requestId: id
            });
        });
    },

    // Resolve a pending request (called by C# response handler)
    _resolve(id, data) {
        const resolve = this._pending[id];
        if (resolve) {
            delete this._pending[id];
            resolve(data);
        }
    },

    // Initialize listener
    init() {
        if (!window.chrome || !window.chrome.webview) return;
        window.chrome.webview.addEventListener('message', (e) => {
            const msg = e.data;
            if (!msg || !msg.action) return;

            // Check if this is a response to a pending request
            if (msg._requestId && this._pending[msg._requestId]) {
                this._resolve(msg._requestId, msg.data);
                return;
            }

            // Dispatch to registered handler
            const handler = this._handlers[msg.action];
            if (handler) handler(msg.data);
        });
    }
};

Bridge.init();
