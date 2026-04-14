// viewer.js - Manifest viewer with AG Grid, tree, detail/preview
const Viewer = {
    profiles: [],
    currentType: 'mail',
    outputRoot: '',
    manifestRows: [],
    gridApi: null,
    selectedFolder: '',

    // --- Initialization ---

    logData: [],
    viewerLastOpened: '',
    _advancedFilterOpen: false,
    _searchTimer: null,
    _currentSearch: '',

    init() {
        Bridge.on('profilesLoaded', (data) => this.onProfilesLoaded(data));
        Bridge.on('manifestLoaded', (data) => this.onManifestLoaded(data));
        Bridge.on('mailBodyLoaded', (data) => this.onMailBodyLoaded(data));
        Bridge.on('attachmentsLoaded', (data) => this.onAttachmentsLoaded(data));
        Bridge.on('filePreview', (data) => this.onFilePreview(data));
        Bridge.on('logLoaded', (data) => this.onLogLoaded(data));
        Bridge.on('configValue', (data) => this.onConfigValue(data));
        Bridge.send('getProfiles');
        Bridge.send('getConfigValue', { key: 'viewer_last_opened' });
        document.addEventListener('click', () => this.hideTreeContextMenu());
    },

    onLogLoaded(data) {
        this.logData = data.rows || [];
        // Refresh grid to show new-item markers
        if (this.gridApi) this.gridApi.refreshCells({ force: true });
    },

    onConfigValue(data) {
        if (data.key === 'viewer_last_opened') {
            this.viewerLastOpened = data.value || '';
            // Update timestamp for next open
            Bridge.send('setConfigValue', { key: 'viewer_last_opened', value: new Date().toISOString() });
        }
    },

    isNewItem(receivedAt) {
        if (!this.viewerLastOpened || !receivedAt) return false;
        return receivedAt > this.viewerLastOpened;
    },

    toggleAdvancedFilter() {
        this._advancedFilterOpen = !this._advancedFilterOpen;
        document.getElementById('advancedFilter').style.display =
            this._advancedFilterOpen ? '' : 'none';
        const btn = document.getElementById('filterToggle');
        btn.innerHTML = (this._advancedFilterOpen ? '&#x25B2;' : '&#x25BC;') + ' Filter';
    },

    // --- Profiles ---

    onProfilesLoaded(data) {
        this.profiles = data.profiles || [];
        const sel = document.getElementById('profileSelect');
        sel.innerHTML = '';
        this.profiles.forEach((p, i) => {
            const opt = document.createElement('option');
            opt.value = String(i);
            opt.textContent = p.name + ' (' + p.type + ')';
            sel.appendChild(opt);
        });
        if (this.profiles.length > 0) {
            sel.value = '0';
            this.loadProfile(0);
        }
    },

    onProfileChange() {
        const sel = document.getElementById('profileSelect');
        const idx = parseInt(sel.value || '0');
        this.loadProfile(idx);
    },

    refresh() {
        const sel = document.getElementById('profileSelect');
        const idx = parseInt(sel.value || '0');
        this.loadProfile(idx);
    },

    loadProfile(idx) {
        const profile = this.profiles[idx];
        if (!profile) return;
        this.currentType = profile.type || 'mail';
        this.selectedFolder = '';
        this._currentSearch = '';
        // Reset search bar and advanced filters
        document.getElementById('searchBox').value = '';
        const df = document.getElementById('filterDateFrom');
        const dt = document.getElementById('filterDateTo');
        const fs = document.getElementById('filterSender');
        const fa = document.getElementById('filterHasAttach');
        if (df) df.value = '';
        if (dt) dt.value = '';
        if (fs) fs.value = '';
        if (fa) fa.checked = false;
        this.clearDetail();
        Bridge.send('getManifest', { profileIndex: idx });
    },

    // --- Manifest data loaded ---

    onManifestLoaded(data) {
        this.manifestRows = data.rows || [];
        this.currentType = data.type || 'mail';
        this.outputRoot = (data.outputRoot || '').replace(/[\\/]+$/, '');
        this.clearDetail();
        this.buildGrid();
        this.buildTree(this.outputRoot);
        this.updateStatus();
        Bridge.send('getLog', { outputRoot: this.outputRoot });
    },

    // --- AG Grid (unified for both types) ---

    buildGrid() {
        // Destroy old grid FIRST, then clear container
        if (this.gridApi) {
            try { this.gridApi.destroy(); } catch(e) {}
            this.gridApi = null;
        }
        const container = document.getElementById('gridContainer');
        container.innerHTML = '';

        const isMail = this.currentType === 'mail';

        const colDefs = isMail ? [
            { field: 'date', headerName: 'Date/Time', maxWidth: 160, sort: 'desc',
              cellRenderer: (params) => {
                  const val = params.value || '';
                  const raw = this.manifestRows[params.data._idx];
                  const isNew = raw && this.isNewItem(raw.received_at || '');
                  return isNew ? '<span style="color:#2563eb;" title="New">&#x25CF;</span> ' + this.esc(val) : this.esc(val);
              }
            },
            { field: 'sender', headerName: 'From', maxWidth: 200 },
            { field: 'subject', headerName: 'Subject', maxWidth: 400 },
            { field: 'attachCount', headerName: '#', maxWidth: 60 }
        ] : [
            { field: 'name', headerName: 'Name', maxWidth: 300,
              cellRenderer: (params) => {
                  const icon = this.fileIcon(params.value);
                  return '<span style="margin-right:4px;">' + icon + '</span>' + this.esc(params.value);
              }
            },
            { field: 'relativePath', headerName: 'Path', maxWidth: 300 },
            { field: 'size', headerName: 'Size', maxWidth: 100,
              valueFormatter: (p) => this.formatSize(p.value) },
            { field: 'modified', headerName: 'Modified', maxWidth: 160 }
        ];

        const rowData = this.manifestRows.map((r, i) => {
            r._idx = i;
            if (isMail) {
                const attPaths = r.attachment_paths || '';
                return {
                    _idx: i,
                    date: (r.received_at || '').replace('T', ' '),
                    sender: r.sender_name || r.sender_email || '',
                    subject: r.subject || '',
                    attachCount: attPaths ? attPaths.split('|').length : 0,
                    _folderPath: (r.folder_path || '').replace(/^[\\/]+/, '').replace(/\//g, '\\')
                };
            } else {
                let fp = r.folder_path || '';
                if (this.outputRoot && fp.toLowerCase().startsWith(this.outputRoot.toLowerCase()))
                    fp = fp.substring(this.outputRoot.length);
                fp = fp.replace(/^[\\/]+/, '').replace(/\//g, '\\') || '.';
                return {
                    _idx: i,
                    name: r.file_name || '',
                    relativePath: r.relative_path || '',
                    size: parseInt(r.file_size) || 0,
                    modified: r.modified_at || '',
                    _folderPath: fp
                };
            }
        });

        const gridOptions = {
            columnDefs: colDefs,
            rowData: rowData,
            rowSelection: 'single',
            onSelectionChanged: () => this.onRowSelected(),
            defaultColDef: { resizable: true, sortable: true, filter: true },
            animateRows: false,
            headerHeight: 34,
            rowHeight: 30,
            isExternalFilterPresent: () => this.hasFilter(),
            doesExternalFilterPass: (node) => this.passesFilter(node),
            onGridReady: () => { this.updateStatus(); },
            onFirstDataRendered: (params) => {
                this.fitColumns(params.api);
                const firstNode = params.api.getDisplayedRowAtIndex(0);
                if (firstNode) firstNode.setSelected(true);
            },
            // Keyboard navigation: ensure selection follows focus
            navigateToNextCell: (params) => {
                const nextCell = params.nextCellPosition;
                if (nextCell) {
                    const node = this.gridApi.getDisplayedRowAtIndex(nextCell.rowIndex);
                    if (node) node.setSelected(true);
                }
                return nextCell;
            }
        };

        try {
            this.gridApi = agGrid.createGrid(container, gridOptions);
        } catch(e) {
            container.innerHTML = '<div style="padding:12px;color:red;">Grid error: ' + e.message + '</div>';
        }
    },

    // --- Tree ---

    // Folder path counts (stored for tree filtering)
    _folderCounts: {},

    buildTree(outputRoot) {
        const container = document.getElementById('treeContainer');
        container.innerHTML = '';

        this._folderCounts = {};
        this._folderLastModified = {};
        this.manifestRows.forEach(r => {
            let fp = r.folder_path || '';
            if (this.currentType !== 'mail' && outputRoot &&
                fp.toLowerCase().startsWith(outputRoot.toLowerCase()))
                fp = fp.substring(outputRoot.length);
            fp = fp.replace(/^[\\/]+/, '').replace(/\//g, '\\') || '.';
            this._folderCounts[fp] = (this._folderCounts[fp] || 0) + 1;
            // Track latest datetime per folder
            const ts = this.currentType === 'mail' ? (r.received_at || '') : (r.modified_at || '');
            if (ts && (!this._folderLastModified[fp] || ts > this._folderLastModified[fp]))
                this._folderLastModified[fp] = ts;
        });

        this._buildTreeNodes(container, this._folderCounts);
    },

    _buildTreeNodes(container, folderCounts) {
        const sorted = Object.keys(folderCounts).sort(
            (a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));

        sorted.forEach(fp => {
            const parts = fp === '.' ? ['.'] : fp.split(/[\\/]/);
            let parentEl = container;
            let cumPath = '';

            parts.forEach((part, i) => {
                cumPath = cumPath ? cumPath + '\\' + part : part;
                const existing = container.querySelector(
                    `[data-path="${CSS.escape(cumPath)}"]`);
                if (existing) {
                    parentEl = existing.querySelector('.tree-children') || existing;
                    return;
                }

                const node = document.createElement('div');
                node.dataset.path = cumPath;

                const label = document.createElement('div');
                label.className = 'tree-node';
                const isLeaf = i === parts.length - 1;
                const count = isLeaf ? folderCounts[fp] : 0;

                const chevron = document.createElement('span');
                chevron.className = 'tree-chevron';
                chevron.innerHTML = '&#x25BC;';
                chevron.onclick = (e) => { e.stopPropagation(); this.toggleTreeNode(node); };

                const text = document.createElement('span');
                text.style.cssText = 'flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;';
                text.textContent = part;

                label.appendChild(chevron);
                label.appendChild(text);
                if (count > 0) {
                    const badge = document.createElement('span');
                    badge.className = 'tree-count';
                    badge.textContent = count;
                    label.appendChild(badge);
                }
                if (isLeaf && this._folderLastModified && this._folderLastModified[fp]) {
                    const ts = document.createElement('span');
                    ts.className = 'tree-date';
                    ts.textContent = this._folderLastModified[fp].replace('T', ' ');
                    label.appendChild(ts);
                }
                label.onclick = () => this.onTreeSelect(cumPath, label);
                label.addEventListener('contextmenu', (e) => {
                    e.preventDefault();
                    this.showTreeContextMenu(e.clientX, e.clientY, cumPath);
                });
                node.appendChild(label);

                const children = document.createElement('div');
                children.className = 'tree-children';
                node.appendChild(children);

                parentEl.appendChild(node);
                parentEl = children;
            });
        });
    },

    toggleTreeNode(node) {
        const children = node.querySelector('.tree-children');
        const chevron = node.querySelector('.tree-chevron');
        if (!children) return;
        const collapsed = children.style.display === 'none';
        children.style.display = collapsed ? '' : 'none';
        chevron.innerHTML = collapsed ? '&#x25BC;' : '&#x25B6;';
    },

    expandAllTree() {
        document.querySelectorAll('#treeContainer .tree-children').forEach(
            el => el.style.display = '');
        document.querySelectorAll('#treeContainer .tree-chevron').forEach(
            el => el.innerHTML = '&#x25BC;');
    },

    collapseAllTree() {
        document.querySelectorAll('#treeContainer .tree-children').forEach(
            el => el.style.display = 'none');
        document.querySelectorAll('#treeContainer .tree-chevron').forEach(
            el => el.innerHTML = '&#x25B6;');
    },

    onTreeSelect(path, labelEl) {
        document.querySelectorAll('.tree-node.selected').forEach(
            el => el.classList.remove('selected'));
        labelEl.classList.add('selected');
        this.selectedFolder = path === '.' ? '' : path;
        this.applyFilter();
    },

    showTreeContextMenu(x, y, folderPath) {
        // For mail profiles, tree folder paths are Outlook paths (not filesystem paths),
        // so "Open in Explorer" is only meaningful for folder profiles.
        if (this.currentType !== 'folder') return;
        const menu = document.getElementById('contextMenu');
        menu.innerHTML = '';
        const item = document.createElement('div');
        item.className = 'ctx-item';
        item.textContent = 'Open in Explorer';
        item.onclick = () => {
            this.hideTreeContextMenu();
            const base = this.outputRoot || '';
            const fullPath = base ? base + '\\' + folderPath.replace(/\//g, '\\')
                                  : folderPath.replace(/\//g, '\\');
            if (fullPath) Bridge.send('openDirectory', { path: fullPath });
        };
        menu.appendChild(item);
        menu.style.display = 'block';
        menu.style.left = Math.min(x, window.innerWidth - 180) + 'px';
        menu.style.top = Math.min(y, window.innerHeight - 40) + 'px';
    },

    hideTreeContextMenu() {
        document.getElementById('contextMenu').style.display = 'none';
    },

    // Re-filter tree to show only folders containing search-matched items
    filterTree() {
        const container = document.getElementById('treeContainer');
        if (!this.gridApi) return;

        const q = this._currentSearch;
        if (!q) {
            // No search: show all tree nodes
            container.querySelectorAll('[data-path]').forEach(n => n.style.display = '');
            return;
        }

        // Collect visible folder paths from filtered grid rows
        const visibleFolders = new Set();
        this.gridApi.forEachNodeAfterFilter(node => {
            const fp = (node.data._folderPath || '').toLowerCase();
            if (fp) {
                visibleFolders.add(fp);
                // Also add all parent paths
                const parts = fp.split(/[\\/]/);
                let cum = '';
                for (const part of parts) {
                    cum = cum ? cum + '\\' + part : part;
                    visibleFolders.add(cum);
                }
            }
        });

        container.querySelectorAll('[data-path]').forEach(n => {
            const path = (n.dataset.path || '').toLowerCase();
            n.style.display = visibleFolders.has(path) ? '' : 'none';
        });
    },

    // --- Search + Filter ---

    onSearchInput() {
        clearTimeout(this._searchTimer);
        this._searchTimer = setTimeout(() => this.onSearch(), 150);
    },

    onSearch() {
        // Search always targets root folder (clear folder selection)
        this.selectedFolder = '';
        document.querySelectorAll('.tree-node.selected').forEach(
            el => el.classList.remove('selected'));
        this.applyFilter();
    },

    hasFilter() {
        if (this._currentSearch.length > 0 || this.selectedFolder.length > 0) return true;
        const df = document.getElementById('filterDateFrom');
        const dt = document.getElementById('filterDateTo');
        const fs = document.getElementById('filterSender');
        const fa = document.getElementById('filterHasAttach');
        if (df && df.value) return true;
        if (dt && dt.value) return true;
        if (fs && fs.value.trim()) return true;
        if (fa && fa.checked) return true;
        return false;
    },

    passesFilter(node) {
        const data = node.data;
        if (this.selectedFolder) {
            const fp = (data._folderPath || '').toLowerCase();
            const sf = this.selectedFolder.toLowerCase();
            if (fp !== sf && !fp.startsWith(sf + '\\') && !fp.startsWith(sf + '/')) return false;
        }
        if (this._currentSearch) {
            const text = Object.values(data).join(' ').toLowerCase();
            if (!text.includes(this._currentSearch)) return false;
        }
        // Advanced filters
        const dateFrom = (document.getElementById('filterDateFrom') || {}).value || '';
        const dateTo = (document.getElementById('filterDateTo') || {}).value || '';
        const senderFilter = ((document.getElementById('filterSender') || {}).value || '').trim().toLowerCase();
        const hasAttach = (document.getElementById('filterHasAttach') || {}).checked || false;
        if (dateFrom && data.date && data.date < dateFrom) return false;
        if (dateTo && data.date && data.date > dateTo) return false;
        if (senderFilter && data.sender && !data.sender.toLowerCase().includes(senderFilter)) return false;
        if (hasAttach && (data.attachCount || 0) === 0) return false;
        return true;
    },

    applyFilter() {
        this._currentSearch = (document.getElementById('searchBox').value || '').trim().toLowerCase();
        if (this.gridApi) this.gridApi.onFilterChanged();
        this.updateStatus();
        this.filterTree();
    },

    updateStatus() {
        let count = 0;
        if (this.gridApi) this.gridApi.forEachNodeAfterFilter(() => count++);
        document.getElementById('statusText').textContent = count + ' items';
    },

    // --- Row selection ---

    onRowSelected() {
        const rows = this.gridApi ? this.gridApi.getSelectedRows() : [];
        if (rows.length === 0) return;

        const idx = rows[0]._idx;
        const r = this.manifestRows[idx];
        if (!r) return;

        if (this.currentType === 'mail') this.showMailDetail(r);
        else this.showFolderDetail(r);
    },

    // --- Detail panel ---

    clearDetail() {
        document.getElementById('emptyState').style.display='';
        document.getElementById('mailDetail').style.display='none';
        document.getElementById('folderDetail').style.display='none';
        document.getElementById('attachPreviewPanel').style.display='none';
        this._attachPreviewActive = false;
    },

    // --- Mail detail ---

    showMailDetail(r) {
        document.getElementById('emptyState').style.display='none';
        document.getElementById('mailDetail').style.display='flex';
        document.getElementById('folderDetail').style.display='none';
        document.getElementById('attachPreviewPanel').style.display='none';
        this._attachPreviewActive = false;

        let headerHtml =
            '<b>From:</b> ' + this.esc(r.sender_name || '') +
            ' &lt;' + this.esc(r.sender_email || '') + '&gt;<br>';
        if (r.to_recipients)
            headerHtml += '<b>To:</b> ' + this.esc(r.to_recipients) + '<br>';
        if (r.cc_recipients)
            headerHtml += '<b>CC:</b> ' + this.esc(r.cc_recipients) + '<br>';
        headerHtml +=
            '<b>Date:</b> ' + this.esc(r.received_at || '') + '<br>' +
            '<b>Subject:</b> ' + this.esc(r.subject || '');
        document.getElementById('mailHeader').innerHTML = headerHtml;
        if (r.mail_folder) {
            const openDirBtn = document.createElement('button');
            openDirBtn.textContent = 'Open folder';
            openDirBtn.style.cssText = 'margin-top:4px;padding:2px 8px;font-size:11px;background:none;border:1px solid #d0d0d0;border-radius:4px;cursor:pointer;';
            openDirBtn.onclick = () => Bridge.send('openDirectory', { path: r.mail_folder });
            document.getElementById('mailHeader').appendChild(openDirBtn);
        }

        Bridge.send('getMailBody', {
            bodyPath: r.body_path || '', bodyText: r.body_text || ''
        });
        Bridge.send('getAttachments', { mailFolder: r.mail_folder || '' });
    },

    onMailBodyLoaded(data) {
        document.getElementById('mailBodyText').textContent = data.body || '';
    },

    onAttachmentsLoaded(data) {
        const el = document.getElementById('attachList');
        el.innerHTML = '';
        (data.files || []).forEach(f => {
            const div = document.createElement('div');
            div.className = 'attach-item';
            const icon = this.fileIcon(f.name);
            div.innerHTML = '<span style="margin-right:4px;">' + icon + '</span>' + this.esc(f.name);
            div.onclick = () => this.previewAttachment(f.name, f.path);
            div.ondblclick = () => Bridge.send('openFile', { path: f.path });
            el.appendChild(div);
        });
    },

    _previewedAttachPath: null,

    previewAttachment(name, path) {
        this._previewedAttachPath = path;
        // Hide mail detail, show full-panel preview
        document.getElementById('mailDetail').style.display='none';
        const panel = document.getElementById('attachPreviewPanel');
        panel.style.display='flex';
        document.getElementById('attachPreviewName').textContent = name;

        const container = document.getElementById('attachPreviewContainer');
        container.innerHTML = '<div class="preview-placeholder">Loading...</div>';

        const ext = (name.split('.').pop() || '').toLowerCase();
        const previewType = this.getPreviewType(ext);
        if (previewType === 'none') {
            container.innerHTML = '<div class="preview-placeholder">No preview for .' +
                this.esc(ext) + '</div>';
            return;
        }
        this._attachPreviewActive = true;
        Bridge.send('getFilePreview', { filePath: path, fileName: name, previewType: previewType });
    },

    closeAttachPreview() {
        document.getElementById('attachPreviewPanel').style.display='none';
        document.getElementById('mailDetail').style.display='flex';
        this._attachPreviewActive = false;
    },

    openPreviewedAttach() {
        if (this._previewedAttachPath)
            Bridge.send('openFile', { path: this._previewedAttachPath });
    },

    // --- Folder detail + preview ---

    showFolderDetail(r) {
        document.getElementById('emptyState').style.display='none';
        document.getElementById('mailDetail').style.display='none';
        document.getElementById('folderDetail').style.display='flex';

        const fi = document.getElementById('fileInfo');
        fi.innerHTML = '<b>Name:</b> ' + this.esc(r.file_name || '') +
            ' &nbsp; <b>Size:</b> ' + this.formatSize(parseInt(r.file_size) || 0) +
            ' &nbsp; <b>Modified:</b> ' + this.esc(r.modified_at || '');
        const openBtn = document.createElement('fluent-button');
        openBtn.setAttribute('size', 'small');
        openBtn.textContent = 'Open';
        openBtn.style.marginLeft = '8px';
        openBtn.onclick = () => this.openCurrentFile();
        fi.appendChild(openBtn);
        const openDirBtn = document.createElement('fluent-button');
        openDirBtn.setAttribute('size', 'small');
        openDirBtn.textContent = 'Open folder';
        openDirBtn.style.marginLeft = '4px';
        openDirBtn.onclick = () => Bridge.send('openDirectory', { path: r.file_path || '' });
        fi.appendChild(openDirBtn);

        this._currentFilePath = r.file_path || '';
        this.requestPreview(r.file_path || '', r.file_name || '');
    },

    openCurrentFile() {
        if (this._currentFilePath)
            Bridge.send('openFile', { path: this._currentFilePath });
    },

    // --- File preview ---

    requestPreview(filePath, fileName) {
        const container = document.getElementById('previewContainer');
        container.innerHTML = '';

        if (!filePath) {
            container.innerHTML = '<div class="preview-placeholder">No file selected</div>';
            return;
        }

        const ext = (fileName.split('.').pop() || '').toLowerCase();
        const previewType = this.getPreviewType(ext);

        if (previewType === 'none') {
            container.innerHTML = '<div class="preview-placeholder">No preview for .' +
                this.esc(ext) + ' files</div>';
            return;
        }

        // Show loading
        container.innerHTML = '<div class="preview-placeholder">Loading preview...</div>';

        Bridge.send('getFilePreview', {
            filePath: filePath, fileName: fileName, previewType: previewType
        });
    },

    getPreviewType(ext) {
        if (['pdf'].includes(ext)) return 'pdf';
        if (['png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg', 'webp', 'ico'].includes(ext)) return 'image';
        if (['html', 'htm'].includes(ext)) return 'html';
        if (['md', 'markdown'].includes(ext)) return 'markdown';
        if (['txt', 'csv', 'json', 'xml', 'log', 'ini', 'cfg',
             'yaml', 'yml', 'toml', 'bat', 'ps1', 'sh', 'py', 'js', 'cs', 'css',
             'sql', 'vbs', 'bas', 'cls', 'frm'].includes(ext)) return 'text';
        if (['xlsx', 'xls'].includes(ext)) return 'excel';
        if (['docx'].includes(ext)) return 'docx';
        if (['pptx'].includes(ext)) return 'pptx';
        return 'none';
    },

    onFilePreview(data) {
        // Route to attachment preview or folder preview
        const containerId = this._attachPreviewActive
            ? 'attachPreviewContainer' : 'previewContainer';
        const container = document.getElementById(containerId);
        container.innerHTML = '';

        switch (data.type) {
            case 'pdf':
                this.renderPdfPreview(container, data);
                break;
            case 'image':
                this.renderImagePreview(container, data);
                break;
            case 'html':
                this.renderHtmlPreview(container, data);
                break;
            case 'markdown':
                this.renderMarkdownPreview(container, data);
                break;
            case 'text':
                this.renderTextPreview(container, data);
                break;
            case 'excel':
                this.renderExcelPreview(container, data);
                break;
            case 'docx':
                this.renderDocxPreview(container, data);
                break;
            case 'pptx':
                this.renderPptxPreview(container, data);
                break;
            default:
                container.innerHTML = '<div class="preview-placeholder">Unsupported format</div>';
        }
    },

    renderPdfPreview(container, data) {
        if (data.virtualPath) {
            // Show loading spinner until PDF renders
            const loading = document.createElement('div');
            loading.className = 'preview-placeholder';
            loading.textContent = 'Loading PDF...';
            container.appendChild(loading);

            const iframe = document.createElement('iframe');
            iframe.style.display = 'none';
            iframe.onload = () => { loading.remove(); iframe.style.display = ''; };
            iframe.src = data.virtualPath;
            container.appendChild(iframe);
        } else {
            container.innerHTML = '<div class="preview-placeholder">Cannot preview this PDF</div>';
        }
    },

    renderImagePreview(container, data) {
        const img = document.createElement('img');
        img.src = 'data:image/' + (data.ext || 'png') + ';base64,' + data.content;
        img.style.padding = '12px';
        img.onerror = () => {
            container.innerHTML = '<div class="preview-placeholder">Cannot render image</div>';
        };
        container.appendChild(img);
    },

    renderTextPreview(container, data) {
        const pre = document.createElement('pre');
        pre.textContent = data.content || '';
        container.appendChild(pre);
    },

    renderHtmlPreview(container, data) {
        const wrapper = document.createElement('div');
        wrapper.style.cssText = 'flex:1;overflow:auto;';
        // Use sandboxed iframe to render HTML safely
        const iframe = document.createElement('iframe');
        iframe.style.cssText = 'border:none;width:100%;height:100%;';
        iframe.sandbox = 'allow-same-origin';
        container.appendChild(iframe);
        iframe.onload = () => {
            try {
                iframe.contentDocument.open();
                iframe.contentDocument.write(data.content || '');
                iframe.contentDocument.close();
            } catch(e) {}
        };
        // Trigger load
        iframe.src = 'about:blank';
    },

    renderMarkdownPreview(container, data) {
        const wrapper = document.createElement('div');
        wrapper.className = 'prose';
        wrapper.style.cssText = 'flex:1;overflow:auto;padding:16px;font-size:13px;line-height:1.7;';
        try {
            wrapper.innerHTML = marked.parse(data.content || '');
        } catch(e) {
            wrapper.textContent = data.content || '';
        }
        container.appendChild(wrapper);
    },

    renderExcelPreview(container, data) {
        try {
            const binary = atob(data.content);
            const bytes = new Uint8Array(binary.length);
            for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

            const workbook = XLSX.read(bytes, { type: 'array' });

            // Tab bar for multiple sheets
            if (workbook.SheetNames.length > 1) {
                const tabBar = document.createElement('div');
                tabBar.style.cssText = 'display:flex;gap:0;border-bottom:1px solid oklch(var(--bc) / 0.15);flex-shrink:0;overflow-x:auto;';
                container.appendChild(tabBar);

                const contentArea = document.createElement('div');
                contentArea.style.cssText = 'flex:1;overflow:auto;padding:8px;';
                container.appendChild(contentArea);

                const showSheet = (name) => {
                    const sheet = workbook.Sheets[name];
                    contentArea.innerHTML = XLSX.utils.sheet_to_html(sheet, { editable: false });
                    tabBar.querySelectorAll('.sheet-tab').forEach(t => {
                        t.style.borderBottom = t.dataset.name === name ? '2px solid oklch(var(--p))' : '2px solid transparent';
                        t.style.color = t.dataset.name === name ? 'oklch(var(--p))' : '';
                    });
                };

                workbook.SheetNames.forEach((name, i) => {
                    const tab = document.createElement('button');
                    tab.className = 'sheet-tab';
                    tab.dataset.name = name;
                    tab.textContent = name;
                    tab.style.cssText = 'padding:6px 14px;font-size:12px;background:none;border:none;' +
                        'border-bottom:2px solid transparent;cursor:pointer;white-space:nowrap;';
                    tab.onclick = () => showSheet(name);
                    tabBar.appendChild(tab);
                });

                showSheet(workbook.SheetNames[0]);
            } else {
                // Single sheet, no tabs
                const wrapper = document.createElement('div');
                wrapper.style.cssText = 'flex:1;overflow:auto;padding:8px;';
                wrapper.innerHTML = XLSX.utils.sheet_to_html(
                    workbook.Sheets[workbook.SheetNames[0]], { editable: false });
                container.appendChild(wrapper);
            }
        } catch (e) {
            container.innerHTML = '<div class="preview-placeholder">Failed to parse Excel file</div>';
        }
    },

    renderDocxPreview(container, data) {
        try {
            const binary = atob(data.content);
            const buf = new ArrayBuffer(binary.length);
            const view = new Uint8Array(buf);
            for (let i = 0; i < binary.length; i++) view[i] = binary.charCodeAt(i);

            const wrapper = document.createElement('div');
            wrapper.style.cssText = 'flex:1;overflow:auto;padding:12px;';
            container.appendChild(wrapper);

            docx.renderAsync(buf, wrapper, null, {
                ignoreWidth: true,
                ignoreHeight: true
            }).catch((err) => {
                wrapper.innerHTML = '<div class="preview-placeholder">Failed to render: ' +
                    (err && err.message ? err.message : 'unknown error') + '</div>';
            });
        } catch (e) {
            container.innerHTML = '<div class="preview-placeholder">Failed to parse: ' +
                e.message + '</div>';
        }
    },

    renderPptxPreview(container, data) {
        try {
            const binary = atob(data.content);
            const bytes = new Uint8Array(binary.length);
            for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

            const wrapper = document.createElement('div');
            wrapper.style.cssText = 'flex:1;overflow:auto;padding:16px;';
            container.appendChild(wrapper);

            // pptx is a ZIP containing XML slides
            JSZip.loadAsync(bytes).then(zip => {
                const slideFiles = Object.keys(zip.files)
                    .filter(f => f.match(/^ppt\/slides\/slide\d+\.xml$/))
                    .sort();

                if (slideFiles.length === 0) {
                    wrapper.innerHTML = '<div class="preview-placeholder">No slides found</div>';
                    return;
                }

                const promises = slideFiles.map(f => zip.file(f).async('text'));
                Promise.all(promises).then(xmlTexts => {
                    xmlTexts.forEach((xml, i) => {
                        const slide = document.createElement('div');
                        slide.style.cssText = 'border:1px solid oklch(var(--bc) / 0.15);' +
                            'border-radius:8px;padding:20px;margin-bottom:12px;background:white;';

                        const header = document.createElement('div');
                        header.style.cssText = 'font-size:11px;color:oklch(var(--bc) / 0.4);margin-bottom:8px;';
                        header.textContent = 'Slide ' + (i + 1);
                        slide.appendChild(header);

                        // Extract text from XML
                        const texts = [];
                        const parser = new DOMParser();
                        const doc = parser.parseFromString(xml, 'text/xml');
                        const tElements = doc.getElementsByTagNameNS(
                            'http://schemas.openxmlformats.org/drawingml/2006/main', 't');
                        let lastParaId = null;
                        for (let t of tElements) {
                            // Group by paragraph (a:p parent)
                            let p = t.parentNode;
                            while (p && p.localName !== 'p') p = p.parentNode;
                            const paraId = p ? Array.from(p.parentNode.children).indexOf(p) : -1;
                            if (lastParaId !== null && paraId !== lastParaId) texts.push('\n');
                            texts.push(t.textContent);
                            lastParaId = paraId;
                        }

                        const content = document.createElement('div');
                        content.style.cssText = 'font-size:13px;line-height:1.6;white-space:pre-wrap;';
                        content.textContent = texts.join('');
                        slide.appendChild(content);

                        wrapper.appendChild(slide);
                    });
                });
            }).catch(() => {
                wrapper.innerHTML = '<div class="preview-placeholder">Failed to parse PowerPoint file</div>';
            });
        } catch (e) {
            container.innerHTML = '<div class="preview-placeholder">Failed to parse PowerPoint file</div>';
        }
    },

    // --- File type icons ---

    fileIcon(fileName) {
        const ext = (fileName || '').split('.').pop().toLowerCase();
        // Inline SVG icons with distinct colors per file type
        const doc = 'M3 1h7l4 4v10H3V1z'; // document shape
        const svg = (color) =>
            '<svg width="14" height="14" viewBox="0 0 16 16" style="vertical-align:middle;">' +
            '<path d="' + doc + '" fill="' + color + '" opacity="0.85"/>' +
            '<path d="M10 1v4h4" fill="none" stroke="' + color + '" stroke-width="0.8" opacity="0.5"/>' +
            '</svg>';
        const colorMap = {
            pdf: '#DC2626',
            doc: '#2563EB', docx: '#2563EB',
            xls: '#16A34A', xlsx: '#16A34A',
            ppt: '#EA580C', pptx: '#EA580C',
            png: '#9333EA', jpg: '#9333EA', jpeg: '#9333EA',
            gif: '#9333EA', bmp: '#9333EA', svg: '#9333EA',
            webp: '#9333EA', ico: '#9333EA',
            zip: '#D97706', '7z': '#D97706', rar: '#D97706',
            gz: '#D97706', tar: '#D97706',
            txt: '#6B7280', log: '#6B7280', csv: '#6B7280',
            html: '#2563EB', htm: '#2563EB',
            js: '#0D9488', cs: '#0D9488', py: '#0D9488', ps1: '#0D9488',
            bat: '#0D9488', sh: '#0D9488', vbs: '#0D9488',
            bas: '#0D9488', cls: '#0D9488', frm: '#0D9488',
            json: '#0D9488', xml: '#0D9488', yaml: '#0D9488',
            yml: '#0D9488', ini: '#0D9488', cfg: '#0D9488',
            toml: '#0D9488',
            msg: '#4F46E5', eml: '#4F46E5',
            mp4: '#DB2777', avi: '#DB2777', mkv: '#DB2777',
            mov: '#DB2777', wmv: '#DB2777',
            mp3: '#0891B2', wav: '#0891B2', flac: '#0891B2',
            m4a: '#0891B2', wma: '#0891B2',
        };
        return svg(colorMap[ext] || '#9CA3AF');
    },

    // --- Helpers ---

    fitColumns(api) {
        if (!api) api = this.gridApi;
        if (!api) return;
        // Always fill available width so columns follow resize
        api.sizeColumnsToFit();
    },

    formatSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1048576) return Math.round(bytes / 1024) + ' KB';
        return (bytes / 1048576).toFixed(1) + ' MB';
    },

    esc(s) {
        const d = document.createElement('div');
        d.textContent = s;
        return d.innerHTML;
    }
};

// Init after a short delay to ensure Fluent Web Components are registered
setTimeout(() => Viewer.init(), 100);
