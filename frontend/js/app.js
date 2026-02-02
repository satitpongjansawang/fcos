// FCOS Application
class FCOSApp {
    constructor() {
        this.currentRevision = null;
        this.init();
    }

    init() {
        this.bindElements();
        this.bindEvents();
        this.loadRevisions();
    }

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
        // Upload area events
        this.uploadArea.addEventListener('click', () => this.fileInput.click());
        this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));

        // Drag and drop
        this.uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            this.uploadArea.classList.add('dragover');
        });

        this.uploadArea.addEventListener('dragleave', () => {
            this.uploadArea.classList.remove('dragover');
        });

        this.uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            this.uploadArea.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                this.uploadFile(files[0]);
            }
        });

        // Export buttons
        this.btnIssueDO.addEventListener('click', () => this.exportReport('issue-do'));
        this.btnDeliveryDaily.addEventListener('click', () => this.exportReport('delivery-daily'));
    }

    handleFileSelect(e) {
        const file = e.target.files[0];
        if (file) {
            this.uploadFile(file);
        }
    }

    async uploadFile(file) {
        // Validate file
        const validTypes = ['.xlsx', '.xls'];
        const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
        
        if (!validTypes.includes(ext)) {
            this.showToast('กรุณาเลือกไฟล์ Excel (.xlsx หรือ .xls)', 'error');
            return;
        }

        if (file.size > 10 * 1024 * 1024) {
            this.showToast('ขนาดไฟล์ต้องไม่เกิน 10MB', 'error');
            return;
        }

        // Show progress
        this.uploadProgress.style.display = 'block';
        this.progressFill.style.width = '0%';
        this.progressText.textContent = 'กำลังอัพโหลด...';

        const formData = new FormData();
        formData.append('file', file);

        try {
            const response = await fetch('/api/upload', {
                method: 'POST',
                body: formData
            });

            // Simulate progress
            let progress = 0;
            const progressInterval = setInterval(() => {
                progress += 10;
                if (progress <= 90) {
                    this.progressFill.style.width = progress + '%';
                }
            }, 100);

            const result = await response.json();
            
            clearInterval(progressInterval);
            this.progressFill.style.width = '100%';
            this.progressText.textContent = 'อัพโหลดสำเร็จ!';

            if (result.success) {
                this.showToast('อัพโหลดไฟล์สำเร็จ!', 'success');
                this.loadRevisions();
                this.selectRevision(result.revision);
                
                // Reset upload area
                setTimeout(() => {
                    this.uploadProgress.style.display = 'none';
                    this.fileInput.value = '';
                }, 1500);
            } else {
                throw new Error(result.error || 'Upload failed');
            }

        } catch (error) {
            console.error('Upload error:', error);
            this.progressText.textContent = 'อัพโหลดไม่สำเร็จ';
            this.showToast(error.message || 'เกิดข้อผิดพลาดในการอัพโหลด', 'error');
            
            setTimeout(() => {
                this.uploadProgress.style.display = 'none';
            }, 2000);
        }
    }

    async loadRevisions() {
        try {
            const response = await fetch('/api/revisions');
            const result = await response.json();

            if (result.success) {
                this.renderRevisions(result.revisions);
            }
        } catch (error) {
            console.error('Load revisions error:', error);
        }
    }

    renderRevisions(revisions) {
        if (revisions.length === 0) {
            this.revisionList.innerHTML = '<p class="empty-message">ยังไม่มีประวัติการอัพโหลด</p>';
            return;
        }

        this.revisionList.innerHTML = revisions.map(rev => `
            <div class="revision-item ${this.currentRevision?.id === rev.id ? 'active' : ''}" data-id="${rev.id}">
                <div class="revision-info">
                    <h4>📁 ${rev.originalName}</h4>
                    <p>อัพโหลดเมื่อ: ${this.formatDate(rev.uploadDate)}</p>
                    <div class="meta">
                        <span>📊 ${rev.recordCount} รายการ</span>
                        <span>📅 ส่ง: ${rev.deliveryDate || '-'}</span>
                        <span>👥 ${rev.customers?.length || 0} ลูกค้า</span>
                    </div>
                </div>
                <div class="revision-actions">
                    <button class="btn-small btn-select" onclick="app.selectRevision('${rev.id}')">เลือก</button>
                    <button class="btn-small btn-delete" onclick="app.deleteRevision('${rev.id}')">ลบ</button>
                </div>
            </div>
        `).join('');
    }

    selectRevision(revisionOrId) {
        let revision = revisionOrId;
        
        // If it's an ID string, find the revision
        if (typeof revisionOrId === 'string') {
            fetch(`/api/revisions/${revisionOrId}`)
                .then(res => res.json())
                .then(result => {
                    if (result.success) {
                        this.setCurrentRevision(result.revision);
                    }
                });
        } else {
            this.setCurrentRevision(revision);
        }
    }

    setCurrentRevision(revision) {
        this.currentRevision = revision;
        
        // Update UI
        document.getElementById('currentFileName').textContent = revision.originalName;
        document.getElementById('currentUploadDate').textContent = this.formatDate(revision.uploadDate);
        document.getElementById('currentRecordCount').textContent = `${revision.recordCount} รายการ`;
        document.getElementById('currentDeliveryDate').textContent = revision.deliveryDate || '-';
        document.getElementById('currentCustomers').textContent = revision.customers?.join(', ') || '-';

        // Show file info section
        this.fileInfoSection.style.display = 'block';

        // Update revision list active state
        document.querySelectorAll('.revision-item').forEach(item => {
            item.classList.remove('active');
            if (item.dataset.id === revision.id) {
                item.classList.add('active');
            }
        });

        // Scroll to file info
        this.fileInfoSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    async deleteRevision(id) {
        if (!confirm('ต้องการลบ revision นี้ใช่หรือไม่?')) {
            return;
        }

        try {
            const response = await fetch(`/api/revisions/${id}`, {
                method: 'DELETE'
            });
            const result = await response.json();

            if (result.success) {
                this.showToast('ลบ revision สำเร็จ', 'success');
                
                // Clear current if deleted
                if (this.currentRevision?.id === id) {
                    this.currentRevision = null;
                    this.fileInfoSection.style.display = 'none';
                }
                
                this.loadRevisions();
            } else {
                throw new Error(result.error);
            }
        } catch (error) {
            this.showToast('เกิดข้อผิดพลาดในการลบ', 'error');
        }
    }

    async exportReport(type) {
        if (!this.currentRevision) {
            this.showToast('กรุณาเลือกไฟล์ก่อน', 'error');
            return;
        }

        this.showLoading(true);

        try {
            const response = await fetch(`/api/export/${type}/${this.currentRevision.id}`);
            
            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.message || 'Export failed');
            }

            // Get filename from header or generate one
            const contentDisposition = response.headers.get('Content-Disposition');
            let filename = type === 'issue-do' ? 'Issue_DO_Report.xlsx' : 'Delivery_Daily_Report.xlsx';
            if (contentDisposition) {
                const match = contentDisposition.match(/filename="(.+)"/);
                if (match) filename = match[1];
            }

            // Download file
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            a.remove();

            this.showToast(`ดาวน์โหลด ${type === 'issue-do' ? 'Issue D/O' : 'Delivery Daily Report'} สำเร็จ!`, 'success');

        } catch (error) {
            console.error('Export error:', error);
            this.showToast(error.message || 'เกิดข้อผิดพลาดในการ export', 'error');
        } finally {
            this.showLoading(false);
        }
    }

    showToast(message, type = 'success') {
        this.toast.className = `toast ${type} show`;
        this.toast.querySelector('.toast-icon').textContent = type === 'success' ? '✅' : '❌';
        this.toast.querySelector('.toast-message').textContent = message;

        setTimeout(() => {
            this.toast.classList.remove('show');
        }, 3000);
    }

    showLoading(show) {
        this.loadingOverlay.style.display = show ? 'flex' : 'none';
    }

    formatDate(dateString) {
        const date = new Date(dateString);
        return date.toLocaleString('th-TH', {
            year: 'numeric',
            month: 'short',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    }
}

// Initialize app
const app = new FCOSApp();
