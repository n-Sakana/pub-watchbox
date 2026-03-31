// settings.js - Settings form logic
const Settings = {
    profiles: [],
    currentIdx: -1,
    filters: [],
    selectedFilter: -1,

    // --- Initialization ---

    init() {
        Bridge.on('configLoaded', (data) => this.onConfigLoaded(data));
        Bridge.on('accountsLoaded', (data) => this.onAccountsLoaded(data));
        Bridge.on('foldersLoaded', (data) => this.onFoldersLoaded(data));
        Bridge.on('folderSelected', (data) => this.onFolderSelected(data));
        Bridge.on('importResult', (data) => this.onImportResult(data));
        Bridge.on('saveResult', (data) => this.onSaveResult(data));
        Bridge.on('closed', () => window.close());

        Bridge.send('getConfig');
    },

    // --- Config loaded from C# ---

    onConfigLoaded(data) {
        this.profiles = data.profiles || [];
        this.rebuildProfileList();
        if (this.profiles.length > 0) {
            const sel = document.getElementById('profileSelect');
            sel.selectedIndex = Math.min(
                Math.max(this.currentIdx, 0), this.profiles.length - 1);
            this.currentIdx = sel.selectedIndex;
            this.loadProfile(this.currentIdx);
        }
        // Request Outlook accounts
        Bridge.send('getAccounts');
    },

    rebuildProfileList() {
        const sel = document.getElementById('profileSelect');
        sel.innerHTML = '';
        this.profiles.forEach((p, i) => {
            const opt = document.createElement('fluent-option');
            opt.value = String(i);
            opt.textContent = p.name || ('Profile ' + (i + 1));
            sel.appendChild(opt);
        });
    },

    // --- Profile switching ---

    onProfileChange() {
        this.saveToModel();
        const sel = document.getElementById('profileSelect');
        sel.blur(); // close native dropdown
        this.currentIdx = parseInt(sel.value);
        this.loadProfile(this.currentIdx);
        // Reload folders for selected account
        const acct = document.getElementById('fAccount').value;
        Bridge.send('getFolders', { account: acct });
    },

    loadProfile(idx) {
        const p = this.profiles[idx];
        if (!p) return;

        document.getElementById('fName').value = p.name || '';
        document.getElementById('fType').value = p.type || 'mail';
        document.getElementById('fOutputRoot').value = p.output_root || '';
        const fFlat = document.getElementById('fFlat');
        const fShort = document.getElementById('fShortDirname');
        if (p.flat_output === '1') fFlat.setAttribute('checked', ''); else fFlat.removeAttribute('checked');
        if (p.short_dirname === '1') fShort.setAttribute('checked', ''); else fShort.removeAttribute('checked');

        // Mail source
        this.selectOption('fAccount', p.account || '');
        this.selectOption('fOutlookFolder', p.outlook_folder || '');

        // Folder source
        document.getElementById('fSourceFolder').value = p.source_folder || '';
        const fRec = document.getElementById('fRecurse');
        const fUnzip = document.getElementById('fAutoUnzip');
        if (p.recurse !== '0') fRec.setAttribute('checked', ''); else fRec.removeAttribute('checked');
        if (p.auto_unzip === '1') fUnzip.setAttribute('checked', ''); else fUnzip.removeAttribute('checked');

        // Filter
        document.getElementById('fSince').value = p.since || '';
        const mode = p.filter_mode === 'and' ? 'and' : 'or';
        const radioGroup = document.getElementById('filterModeGroup');
        if (radioGroup) radioGroup.value = mode;

        this.filters = (p.filters || '').split(';').filter(k => k.trim().length > 0);
        this.rebuildFilterList();

        // Monitoring
        const setChk = (id, val) => { const el = document.getElementById(id); if (val) el.setAttribute('checked', ''); else el.removeAttribute('checked'); };
        setChk('fNotify', p.notify !== '0');
        setChk('fLog', p.log_enabled !== '0');
        setChk('fManifestHidden', p.manifest_hidden !== '0');

        // Type toggle
        this.onTypeChange();
    },

    saveToModel() {
        if (this.currentIdx < 0 || this.currentIdx >= this.profiles.length) return;
        const p = this.profiles[this.currentIdx];

        p.name = document.getElementById('fName').value;
        p.type = document.getElementById('fType').value;
        p.output_root = document.getElementById('fOutputRoot').value;
        p.flat_output = document.getElementById('fFlat').hasAttribute('checked') ? '1' : '0';
        p.short_dirname = document.getElementById('fShortDirname').hasAttribute('checked') ? '1' : '0';

        p.account = document.getElementById('fAccount').value;
        p.outlook_folder = document.getElementById('fOutlookFolder').value;

        p.source_folder = document.getElementById('fSourceFolder').value;
        p.recurse = document.getElementById('fRecurse').hasAttribute('checked') ? '1' : '0';
        p.auto_unzip = document.getElementById('fAutoUnzip').hasAttribute('checked') ? '1' : '0';

        p.since = document.getElementById('fSince').value;
        const rg = document.getElementById('filterModeGroup');
        p.filter_mode = rg ? (rg.value || 'or') : 'or';
        p.filters = this.filters.join(';');

        const chk = (id) => document.getElementById(id).hasAttribute('checked') ? '1' : '0';
        p.notify = chk('fNotify');
        p.log_enabled = chk('fLog');
        p.manifest_hidden = chk('fManifestHidden');

        // Update combo text
        const sel = document.getElementById('profileSelect');
        if (sel.options[this.currentIdx])
            sel.options[this.currentIdx].textContent = p.name || ('Profile ' + (this.currentIdx + 1));
    },

    // --- Type switching ---

    onTypeChange() {
        const type = document.getElementById('fType').value;
        document.body.className = document.body.className
            .replace(/type-\w+/, '') + ' type-' + type;
        // Short names only for mail
        const sdr = document.getElementById('fShortDirname');
        if (sdr) sdr.style.display = type === 'mail' ? '' : 'none';
    },

    // --- Outlook data ---

    onAccountsLoaded(data) {
        const sel = document.getElementById('fAccount');
        const current = sel.value;
        sel.innerHTML = '<fluent-option value="">(All)</fluent-option>';
        (data.accounts || []).forEach(a => {
            const opt = document.createElement('fluent-option');
            opt.value = a;
            opt.textContent = a;
            sel.appendChild(opt);
        });
        this.selectOption('fAccount', current ||
            (this.profiles[this.currentIdx] || {}).account || '');

        // Load folders for current account
        Bridge.send('getFolders', { account: sel.value });
    },

    onFoldersLoaded(data) {
        const sel = document.getElementById('fOutlookFolder');
        const current = sel.value ||
            (this.profiles[this.currentIdx] || {}).outlook_folder || '';
        sel.innerHTML = '<fluent-option value="">(All)</fluent-option>';
        (data.folders || []).forEach(f => {
            const opt = document.createElement('fluent-option');
            opt.value = f.path;
            opt.textContent = f.display;
            sel.appendChild(opt);
        });
        this.selectOption('fOutlookFolder', current);
    },

    onAccountChange() {
        const acct = document.getElementById('fAccount').value;
        Bridge.send('getFolders', { account: acct });
    },

    // --- Browse folder ---

    browse(field) {
        const current = field === 'output_root'
            ? document.getElementById('fOutputRoot').value
            : document.getElementById('fSourceFolder').value;
        Bridge.send('browseFolder', { field, current });
    },

    onFolderSelected(data) {
        if (!data.path) return;
        if (data.field === 'output_root')
            document.getElementById('fOutputRoot').value = data.path;
        else if (data.field === 'source_folder')
            document.getElementById('fSourceFolder').value = data.path;
    },

    // --- Filter keywords ---

    rebuildFilterList() {
        const el = document.getElementById('filterList');
        el.innerHTML = '';
        this.filters.forEach((kw, i) => {
            const div = document.createElement('div');
            div.className = 'px-2 py-0.5 text-sm cursor-pointer rounded hover:bg-base-200'
                + (i === this.selectedFilter ? ' bg-primary/10 text-primary' : '');
            div.textContent = kw;
            div.onclick = () => { this.selectedFilter = i; this.rebuildFilterList(); };
            el.appendChild(div);
        });
    },

    addFilter() {
        const input = document.getElementById('fNewFilter');
        const kw = input.value.trim();
        if (!kw) return;
        this.filters.push(kw);
        input.value = '';
        input.focus();
        this.rebuildFilterList();
    },

    removeFilter() {
        if (this.selectedFilter < 0 || this.selectedFilter >= this.filters.length) return;
        this.filters.splice(this.selectedFilter, 1);
        this.selectedFilter = -1;
        this.rebuildFilterList();
    },

    // --- Profile management ---

    addProfile() {
        this.saveToModel();
        const newProfile = {
            name: 'Profile ' + (this.profiles.length + 1),
            type: 'mail', output_root: '', account: '', outlook_folder: '',
            source_folder: '', recurse: '1', since: '', filter_mode: 'or',
            filters: '', flat_output: '0', short_dirname: '0',
            notify: '1', log_enabled: '1', manifest_hidden: '1'
        };
        this.profiles.push(newProfile);
        this.rebuildProfileList();
        const sel = document.getElementById('profileSelect');
        sel.selectedIndex = this.profiles.length - 1;
        this.currentIdx = sel.selectedIndex;
        this.loadProfile(this.currentIdx);
    },

    removeProfile() {
        if (this.profiles.length <= 1) {
            showModal('Cannot remove the last profile.', false);
            return;
        }
        const idx = this.currentIdx;
        this.profiles.splice(idx, 1);
        this.currentIdx = Math.min(idx, this.profiles.length - 1);
        this.rebuildProfileList();
        const sel = document.getElementById('profileSelect');
        sel.selectedIndex = this.currentIdx;
        this.loadProfile(this.currentIdx);
    },

    // --- Save / Cancel ---

    save() {
        this.saveToModel();
        Bridge.send('saveConfig', { profiles: this.profiles });
    },

    onSaveResult(data) {
        if (data.ok) Bridge.send('close');
    },

    cancel() {
        Bridge.send('close');
    },

    // --- Import / Export / Reset ---

    importCsv() {
        Bridge.send('importCsv');
    },

    onImportResult(data) {
        if (data.profiles && data.profiles.length > 0) {
            this.saveToModel();
            data.profiles.forEach(p => this.profiles.push(p));
            this.rebuildProfileList();
            const sel = document.getElementById('profileSelect');
            sel.selectedIndex = this.profiles.length - 1;
            this.currentIdx = sel.selectedIndex;
            this.loadProfile(this.currentIdx);
            showModal(data.profiles.length + ' profile(s) imported.', false);
        }
    },

    exportCsv() {
        this.saveToModel();
        Bridge.send('exportCsv', { profiles: this.profiles });
    },

    async resetAll() {
        if (!await showModal('All profiles will be deleted. Continue?', true)) return;
        this.profiles = [{
            name: 'Default', type: 'mail', output_root: '', account: '',
            outlook_folder: '', source_folder: '', recurse: '1', since: '',
            filter_mode: 'or', filters: '', flat_output: '0', short_dirname: '0',
            notify: '1', log_enabled: '1', manifest_hidden: '1'
        }];
        this.currentIdx = 0;
        this.rebuildProfileList();
        document.getElementById('profileSelect').selectedIndex = 0;
        this.loadProfile(0);
    },

    // --- Helpers ---

    selectOption(id, value) {
        const sel = document.getElementById(id);
        // Fluent select: set value directly
        sel.value = value || '';
    }
};

// Start
Settings.init();
