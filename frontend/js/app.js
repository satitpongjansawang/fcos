class FCOSApp {
    constructor() { this.currentRevision = null; this.init(); }
    init() { this.bindElements(); this.bindEvents(); this.loadRevisions(); }
    bindElements() {
        this.uploadArea = document.getElementById('uploadArea');
        this.fileInput = document.getElementById('fileInput');
        this.uploadProgress = document.getElementById('uploadProgress');
        this.progressFill = document.getElementById('progressFill');
        this.progressText = document.getElementById('progressText');
        this.fileInfoSection = document.getElementById('fileInfoSection');
        this.revisionList = document.getElementById('revisionList');
        this.btnIssueDO = document.getElementById('btnIssueDO');
        this.btnDeliveryDaily = document.getElementById('btnDeliveryDaily');
        this.toast = document.getElementById('toast');
        this.loadingOverlay = document.getElementById('loadingOverlay');
    }
    bindEvents() {
        this.uploadArea.addEventListener('click', () => this.fileInput.click());
        this.fileInput.addEventListener('change', (e) => { if (e.target.files[0]) this.uploadFile(e.target.files[0]); });
        this.uploadArea.addEventListener('dragover', (e) => { e.preventDefault(); this.uploadArea.classList.add('dragover'); });
        this.uploadArea.addEventListener('dragleave', () => this.uploadArea.classList.remove('dragover'));
        this.uploadArea.addEventListener('drop', (e) => { e.preventDefault(); this.uploadArea.classList.remove('dragover'); if (e.dataTransfer.files.length) this.uploadFile(e.dataTransfer.files[0]); });
        this.btnIssueDO.addEventListener('click', () => this.exportReport('issue-do'));
        this.btnDeliveryDaily.addEventListener('click', () => this.exportReport('delivery-daily'));
    }
    async uploadFile(file) {
        const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
        if (!['.xlsx', '.xls'].includes(ext)) { this.showToast('กรุณาเลือกไฟล์ .xlsx หรือ .xls', false); return; }
        if (file.size > 10 * 1024 * 1024) { this.showToast('ไฟล์มีขนาดเกิน 10MB', false); return; }
        this.uploadProgress.style.display = 'block'; this.progressFill.style.width = '30%'; this.progressText.textContent = 'กำลังอัพโหลด...';
        try {
            const formData = new FormData(); formData.append('file', file);
            const response = await fetch('/api/upload', { method: 'POST', body: formData });
            this.progressFill.style.width = '100%';
            const result = await response.json();
            if (result.success) { this.currentRevision = result.revision; this.showFileInfo(result.revision); this.loadRevisions(); this.showToast('อัพโหลดสำเร็จ!', true); }
            else { this.showToast(result.error || 'อัพโหลดไม่สำเร็จ', false); }
        } catch (e) { this.showToast('เกิดข้อผิดพลาด: ' + e.message, false); }
        finally { setTimeout(() => { this.uploadProgress.style.display = 'none'; this.progressFill.style.width = '0%'; }, 1000); }
    }
    showFileInfo(revision) {
        this.fileInfoSection.style.display = 'block';
        document.getElementById('infoFileName').textContent = revision.originalName;
        document.getElementById('infoRows').textContent = revision.summary?.totalRows || '-';
        document.getElementById('infoDO').textContent = revision.summary?.doNumbers || '-';
        document.getElementById('infoCustomer').textContent = revision.summary?.customerCodes || '-';
    }
    async loadRevisions() {
        try {
            const res = await fetch('/api/revisions'); const data = await res.json();
            if (data.success && data.revisions.length) {
                this.revisionList.innerHTML = data.revisions.map(r => `
                    <div class="revision-item">
                        <div class="revision-info">
                            <div class="revision-name">${r.originalName}</div>
                            <div class="revision-date">${this.formatDate(r.uploadDate)}</div>
                            <div class="revision-summary">${r.summary?.totalRows || 0} rows | ${r.summary?.doNumbers || 0} DOs</div>
                        </div>
                        <div class="revision-actions">
                            <button onclick="app.selectRevision('${r.id}')">เลือก</button>
                            <button onclick="app.deleteRevision('${r.id}')">ลบ</button>
                        </div>
                    </div>
                `).join('');
                if (!this.currentRevision) { this.currentRevision = data.revisions[0]; this.showFileInfo(data.revisions[0]); }
            } else { this.revisionList.innerHTML = '<p class="empty-text">ยังไม่มีประวัติการอัพโหลด</p>'; }
        } catch (e) { console.error('Load revisions error:', e); }
    }
    async selectRevision(id) {
        try {
            const res = await fetch(`/api/revisions/${id}`); const data = await res.json();
            if (data.success) { this.currentRevision = data.revision; this.showFileInfo(data.revision); this.showToast('เลือก revision สำเร็จ', true); }
        } catch (e) { this.showToast('ไม่สามารถโหลด revision', false); }
    }
    async deleteRevision(id) {
        if (!confirm('ต้องการลบ revision นี้?')) return;
        try {
            const res = await fetch(`/api/revisions/${id}`, { method: 'DELETE' }); const data = await res.json();
            if (data.success) { this.loadRevisions(); this.showToast('ลบสำเร็จ', true); if (this.currentRevision?.id === id) { this.currentRevision = null; this.fileInfoSection.style.display = 'none'; } }
        } catch (e) { this.showToast('ลบไม่สำเร็จ', false); }
    }
    async exportReport(type) {
        if (!this.currentRevision) { this.showToast('กรุณาเลือกไฟล์ก่อน', false); return; }
        this.showLoading(true);
        try {
            const res = await fetch(`/api/export/${type}/${this.currentRevision.id}`);
            if (!res.ok) throw new Error('Export failed');
            const blob = await res.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a'); a.href = url; a.download = `${type}_${new Date().toISOString().split('T')[0]}.xlsx`; a.click();
            window.URL.revokeObjectURL(url); this.showToast('ดาวน์โหลดสำเร็จ!', true);
        } catch (e) { this.showToast('Export ไม่สำเร็จ: ' + e.message, false); }
        finally { this.showLoading(false); }
    }
    showToast(message, success) {
        this.toast.className = `toast show ${success ? 'success' : 'error'}`;
        this.toast.querySelector('.toast-icon').textContent = success ? '✅' : '❌';
        this.toast.querySelector('.toast-message').textContent = message;
        setTimeout(() => this.toast.classList.remove('show'), 3000);
    }
    showLoading(show) { this.loadingOverlay.style.display = show ? 'flex' : 'none'; }
    formatDate(dateString) { return new Date(dateString).toLocaleString('th-TH', { year:'numeric',month:'short',day:'numeric',hour:'2-digit',minute:'2-digit' }); }
}
const app = new FCOSApp();
