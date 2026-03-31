// viewer.js - Manifest viewer with AG Grid, tree, detail/preview
const Viewer = {
    profiles: [],
    currentType: 'mail',
    outputRoot: '',
    manifestRows: [],
    gridApi: null,
    selectedFolder: '',

    // --- Initialization ---

    init() {
        Bridge.on('profilesLoaded', (data) => this.onProfilesLoaded(data));
        Bridge.on('manifestLoaded', (data) => this.onManifestLoaded(data));
        Bridge.on('mailBodyLoaded', (data) => this.onMailBodyLoaded(data));
        Bridge.on('attachmentsLoaded', (data) => this.onAttachmentsLoaded(data));
        Bridge.on('filePreview', (data) => this.onFilePreview(data));
        Bridge.send('getProfiles');
    },

    // --- Profiles ---

    onProfilesLoaded(data) {
        this.profiles = data.profiles || [];
        const sel = document.getElementById('profileSelect');
        sel.innerHTML = '';
        this.profiles.forEach((p, i) => {
            const opt = document.createElement('option');
            opt.value = i;
            opt.textContent = p.name + ' (' + p.type + ')';
            sel.appendChild(opt);
        });
        if (this.profiles.length > 0) {
            sel.selectedIndex = 0;
            this.onProfileChange();
        }
    },

    onProfileChange() {
        const idx = parseInt(document.getElementById('profileSelect').value);
        const profile = this.profiles[idx];
        if (!profile) return;

        this.currentType = profile.type || 'mail';
        this.selectedFolder = '';
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
            { field: 'date', headerName: 'Date', width: 100, sort: 'desc' },
            { field: 'sender', headerName: 'From', width: 140 },
            { field: 'subject', headerName: 'Subject', flex: 1 },
            { field: 'attachCount', headerName: '#', width: 50 }
        ] : [
            { field: 'name', headerName: 'Name', flex: 1 },
            { field: 'relativePath', headerName: 'Path', flex: 1 },
            { field: 'size', headerName: 'Size', width: 90,
              valueFormatter: (p) => this.formatSize(p.value) },
            { field: 'modified', headerName: 'Modified', width: 140 }
        ];

        const rowData = this.manifestRows.map((r, i) => {
            r._idx = i;
            if (isMail) {
                const attPaths = r.attachment_paths || '';
                return {
                    _idx: i,
                    date: (r.received_at || '').substring(0, 10),
                    sender: r.sender_name || r.sender_email || '',
                    subject: r.subject || '',
                    attachCount: attPaths ? attPaths.split('|').length : 0,
                    _folderPath: (r.folder_path || '').replace(/^[\\/]+/, '')
                };
            } else {
                let fp = r.folder_path || '';
                if (this.outputRoot && fp.toLowerCase().startsWith(this.outputRoot.toLowerCase()))
                    fp = fp.substring(this.outputRoot.length);
                fp = fp.replace(/^[\\/]+/, '') || '.';
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
            onFirstDataRendered: (params) => { params.api.sizeColumnsToFit(); }
        };

        try {
            this.gridApi = agGrid.createGrid(container, gridOptions);
        } catch(e) {
            container.innerHTML = '<div style="padding:12px;color:red;">Grid error: ' + e.message + '</div>';
        }
    },

    // --- Tree ---

    buildTree(outputRoot) {
        const container = document.getElementById('treeContainer');
        container.innerHTML = '';

        const folderCounts = {};
        this.manifestRows.forEach(r => {
            let fp = r.folder_path || '';
            if (this.currentType !== 'mail' && outputRoot &&
                fp.toLowerCase().startsWith(outputRoot.toLowerCase()))
                fp = fp.substring(outputRoot.length);
            fp = fp.replace(/^[\\/]+/, '') || '.';
            folderCounts[fp] = (folderCounts[fp] || 0) + 1;
        });

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
                label.innerHTML = '<span class="opacity-40 text-xs">&#x25B6;</span> ' +
                    this.esc(part) +
                    (count > 0 ? ' <span class="opacity-40 ml-1">' + count + '</span>' : '');
                label.onclick = () => this.onTreeSelect(cumPath, label);
                node.appendChild(label);

                const children = document.createElement('div');
                children.className = 'tree-children';
                node.appendChild(children);

                parentEl.appendChild(node);
                parentEl = children;
            });
        });
    },

    onTreeSelect(path, labelEl) {
        document.querySelectorAll('.tree-node.selected').forEach(
            el => el.classList.remove('selected'));
        labelEl.classList.add('selected');
        this.selectedFolder = path === '.' ? '' : path;
        this.applyFilter();
    },

    // --- Search + Filter ---

    onSearch() { this.applyFilter(); },

    hasFilter() {
        const q = (document.getElementById('searchBox').value || '').trim();
        return q.length > 0 || this.selectedFolder.length > 0;
    },

    passesFilter(node) {
        const data = node.data;
        if (this.selectedFolder) {
            const fp = (data._folderPath || '');
            if (!fp.toLowerCase().startsWith(this.selectedFolder.toLowerCase())) return false;
        }
        const q = (document.getElementById('searchBox').value || '').trim().toLowerCase();
        if (q) {
            const text = Object.values(data).join(' ').toLowerCase();
            if (!text.includes(q)) return false;
        }
        return true;
    },

    applyFilter() {
        if (this.gridApi) this.gridApi.onFilterChanged();
        this.updateStatus();
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
        document.getElementById('emptyState').classList.remove('hidden');
        document.getElementById('mailDetail').classList.add('hidden');
        document.getElementById('folderDetail').classList.add('hidden');
    },

    // --- Mail detail ---

    showMailDetail(r) {
        document.getElementById('emptyState').classList.add('hidden');
        document.getElementById('mailDetail').classList.remove('hidden');
        document.getElementById('folderDetail').classList.add('hidden');

        document.getElementById('mailHeader').innerHTML =
            '<b>From:</b> ' + this.esc(r.sender_name || '') +
            ' &lt;' + this.esc(r.sender_email || '') + '&gt;<br>' +
            '<b>Date:</b> ' + this.esc(r.received_at || '') + '<br>' +
            '<b>Subject:</b> ' + this.esc(r.subject || '');

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
            div.textContent = f.name;
            div.ondblclick = () => Bridge.send('openFile', { path: f.path });
            el.appendChild(div);
        });
    },

    // --- Folder detail + preview ---

    showFolderDetail(r) {
        document.getElementById('emptyState').classList.add('hidden');
        document.getElementById('mailDetail').classList.add('hidden');
        document.getElementById('folderDetail').classList.remove('hidden');

        document.getElementById('fileInfo').innerHTML =
            '<b>Name:</b> ' + this.esc(r.file_name || '') +
            ' &nbsp; <b>Size:</b> ' + this.formatSize(parseInt(r.file_size) || 0) +
            ' &nbsp; <b>Modified:</b> ' + this.esc(r.modified_at || '') +
            ' &nbsp; <button class="btn btn-xs btn-ghost ml-2" onclick="Viewer.openCurrentFile()">Open</button>';

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
        const container = document.getElementById('previewContainer');
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
        // WebView2's Chromium has a built-in PDF viewer
        // Use virtual host mapping to serve the file
        if (data.virtualPath) {
            const iframe = document.createElement('iframe');
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
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const html = XLSX.utils.sheet_to_html(sheet, { editable: false });

            const wrapper = document.createElement('div');
            wrapper.style.cssText = 'flex:1;overflow:auto;padding:8px;';
            wrapper.innerHTML = html;
            container.appendChild(wrapper);
        } catch (e) {
            container.innerHTML = '<div class="preview-placeholder">Failed to parse Excel file</div>';
        }
    },

    renderDocxPreview(container, data) {
        try {
            const binary = atob(data.content);
            const bytes = new Uint8Array(binary.length);
            for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

            const wrapper = document.createElement('div');
            wrapper.style.cssText = 'flex:1;overflow:auto;padding:12px;';
            container.appendChild(wrapper);

            docx.renderAsync(bytes, wrapper).catch(() => {
                wrapper.innerHTML = '<div class="preview-placeholder">Failed to render Word document</div>';
            });
        } catch (e) {
            container.innerHTML = '<div class="preview-placeholder">Failed to parse Word file</div>';
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

    // --- Helpers ---

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

Viewer.init();
