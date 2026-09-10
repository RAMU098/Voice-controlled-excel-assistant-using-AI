/**
 * Voice-Controlled Excel Assistant
 * Frontend JavaScript Application
 */

class VoiceExcelApp {
    constructor() {
        this.sessionId = null;
        this.isListening = false;
        this.recognition = null;
        this.commandHistory = [];
        this.assistantTimeout = null;
        this.lastCommand = '';
        this.lastCommandTime = 0;
        this.debounceDelay = 1500; // 1.5 seconds between duplicate commands
        this.isProcessing = false;
        
        this.init();
    }
    
    async init() {
        await this.createSession();
        this.initElements();
        this.initSpeechRecognition();
        this.initEventListeners();
        this.updatePreview();
    }
    
    initElements() {
        // Voice control elements
        this.voiceBtn = document.getElementById('voiceBtn');
        this.commandInput = document.getElementById('commandInput');
        this.sendBtn = document.getElementById('sendBtn');
        this.statusIndicator = document.getElementById('statusIndicator');
        this.historyList = document.getElementById('historyList');
        this.voiceFeedback = document.getElementById('voiceFeedback');
        
        // Excel preview elements
        this.tableHeader = document.getElementById('tableHeader');
        this.tableBody = document.getElementById('tableBody');
        this.sheetTabs = document.getElementById('sheetTabs');
        this.excelPreview = document.getElementById('excelPreview');
        
        // File upload elements
        this.uploadArea = document.getElementById('uploadArea');
        this.uploadContent = document.getElementById('uploadContent');
        this.fileInput = document.getElementById('fileInput');
        
        // Action buttons
        this.actionBtns = document.querySelectorAll('.action-btn');
        this.downloadBtn = document.getElementById('downloadBtn');
        this.helpBtn = document.getElementById('helpBtn');
        
        // Modal elements
        this.helpModal = document.getElementById('helpModal');
        this.closeModal = document.getElementById('closeModal');
        
        // Toast notification
        this.toast = document.getElementById('toast');
        
        // Assistant popup elements
        this.assistantPopup = document.getElementById('assistantPopup');
        this.closeAssistant = document.getElementById('closeAssistant');
        this.assistantStatus = document.getElementById('assistantStatus');
        this.assistantMessage = document.getElementById('assistantMessage');
        this.assistantAction = document.getElementById('assistantAction');
        this.actionText = document.getElementById('actionText');
        this.affectedCells = document.getElementById('affectedCells');
        this.processingIndicator = document.getElementById('processingIndicator');
    }
    
    initSpeechRecognition() {
        // Check for browser support
        const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
        
        if (!SpeechRecognition) {
            console.warn('Speech recognition not supported in this browser');
            this.voiceBtn.disabled = true;
            this.showToast('Voice recognition not supported in your browser. Please use Chrome or Edge.', 'error');
            return;
        }
        
        this.recognition = new SpeechRecognition();
        this.recognition.continuous = true;  // Keep listening
        this.recognition.interimResults = true;  // Get interim results
        this.recognition.lang = 'en-US';
        
        this.speechTimeout = null;
        this.finalTranscript = '';
        this.interimTranscript = '';
        
        this.recognition.onstart = () => {
            this.isListening = true;
            this.voiceBtn.classList.add('listening');
            this.updateStatus('Listening...', 'listening');
            this.finalTranscript = '';
            this.interimTranscript = '';
        };
        
        this.recognition.onresult = (event) => {
            this.interimTranscript = '';
            
            for (let i = event.resultIndex; i < event.results.length; i++) {
                const transcript = event.results[i][0].transcript;
                
                if (event.results[i].isFinal) {
                    this.finalTranscript += transcript + ' ';
                } else {
                    this.interimTranscript += transcript;
                }
            }
            
            // Show what's being said in real-time
            const displayText = this.finalTranscript + this.interimTranscript;
            this.commandInput.value = displayText.trim();
            this.updateStatus('Listening... (speak, then pause)', 'listening');
            
            // Clear any existing timeout
            if (this.speechTimeout) {
                clearTimeout(this.speechTimeout);
            }
            
            // Wait 1.5 seconds of silence before processing
            this.speechTimeout = setTimeout(() => {
                const finalText = this.finalTranscript.trim();
                if (finalText.length > 0) {
                    console.log(`Processing speech after pause: "${finalText}"`);
                    
                    // Apply speech corrections before processing
                    const correctedText = this.correctSpeechErrors(finalText);
                    this.commandInput.value = correctedText;
                    
                    // Stop listening and process
                    this.recognition.stop();
                    this.processCommand(correctedText);
                }
            }, 1500);  // 1.5 second pause before processing
        };
        
        this.recognition.onerror = (event) => {
            console.error('Speech recognition error:', event.error);
            
            if (event.error === 'not-allowed') {
                this.stopListening();
                this.showToast('Microphone access denied. Please allow microphone access and try again.', 'error');
            } else if (event.error === 'no-speech') {
                // Don't stop on no-speech, just update status
                this.updateStatus('No speech detected, keep speaking...', 'listening');
            } else if (event.error !== 'aborted') {
                this.stopListening();
            }
        };
        
        this.recognition.onend = () => {
            // Clear any pending timeout
            if (this.speechTimeout) {
                clearTimeout(this.speechTimeout);
                this.speechTimeout = null;
            }
            this.stopListening();
        };
    }
    
    correctSpeechErrors(text) {
        // Common speech recognition errors
        let corrected = text.toLowerCase();
        
        // "some" → "sum" in Excel contexts
        corrected = corrected.replace(/\bsome\s+of\b/gi, 'sum of');
        corrected = corrected.replace(/\bdo\s+some\b/gi, 'do sum');
        corrected = corrected.replace(/\bthe\s+some\b/gi, 'the sum');
        corrected = corrected.replace(/\bcalculate\s+some\b/gi, 'calculate sum');
        corrected = corrected.replace(/\bfind\s+some\b/gi, 'find sum');
        corrected = corrected.replace(/\bget\s+some\b/gi, 'get sum');
        corrected = corrected.replace(/\badd\s+some\b/gi, 'add sum');
        
        // Common word corrections
        corrected = corrected.replace(/\braw\b/gi, 'row');
        corrected = corrected.replace(/\brows\b/gi, 'row');
        corrected = corrected.replace(/\bcolum\b/gi, 'column');
        corrected = corrected.replace(/\bcaller\b/gi, 'column');
        corrected = corrected.replace(/\bcalm\b/gi, 'column');
        corrected = corrected.replace(/\bbowled\b/gi, 'bold');
        corrected = corrected.replace(/\bweight\b/gi, 'write');
        corrected = corrected.replace(/\bright\b/gi, 'write');
        
        // Sheet corrections
        corrected = corrected.replace(/\bsheet\s+one\b/gi, 'sheet 1');
        corrected = corrected.replace(/\bsheet\s+to\b/gi, 'sheet 2');
        corrected = corrected.replace(/\bsheet\s+too\b/gi, 'sheet 2');
        
        console.log(`Speech correction: "${text}" → "${corrected}"`);
        return corrected;
    }
    
    initEventListeners() {
        // Voice button
        this.voiceBtn?.addEventListener('click', () => this.toggleListening());
        
        // Keyboard shortcut for voice (Space key)
        document.addEventListener('keydown', (e) => {
            if (e.code === 'Space' && e.target === document.body) {
                e.preventDefault();
                this.toggleListening();
            }
        });
        
        // Send button
        this.sendBtn?.addEventListener('click', () => {
            const command = this.commandInput.value.trim();
            if (command) {
                this.processCommand(command);
            }
        });
        
        // Enter key in input
        this.commandInput?.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                const command = this.commandInput.value.trim();
                if (command) {
                    this.processCommand(command);
                }
            }
        });
        
        // Quick action buttons
        this.actionBtns.forEach(btn => {
            btn.addEventListener('click', () => {
                const command = btn.dataset.command;
                if (command) {
                    this.commandInput.value = command;
                    this.processCommand(command);
                }
            });
        });
        
        // File upload - handle click on upload area
        this.uploadContent?.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            if (this.fileInput) {
                this.fileInput.click();
            }
        });
        
        // Also handle click on the upload area itself
        this.uploadArea?.addEventListener('click', (e) => {
            // Only trigger if clicking on the area, not on already handled elements
            if (e.target === this.uploadArea || e.target.closest('.upload-content')) {
                e.preventDefault();
                if (this.fileInput) {
                    this.fileInput.click();
                }
            }
        });
        
        this.fileInput?.addEventListener('change', (e) => {
            if (e.target.files && e.target.files.length > 0) {
                this.uploadFile(e.target.files[0]);
                // Reset the input so the same file can be selected again
                e.target.value = '';
            }
        });
        
        // Drag and drop
        this.uploadArea?.addEventListener('dragover', (e) => {
            e.preventDefault();
            this.uploadContent?.classList.add('dragover');
        });
        
        this.uploadArea?.addEventListener('dragleave', () => {
            this.uploadContent?.classList.remove('dragover');
        });
        
        this.uploadArea?.addEventListener('drop', (e) => {
            e.preventDefault();
            this.uploadContent?.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) {
                this.uploadFile(e.dataTransfer.files[0]);
            }
        });
        
        // Download button
        this.downloadBtn?.addEventListener('click', () => this.downloadFile());
        
        // Help modal
        this.helpBtn?.addEventListener('click', () => this.showHelpModal());
        this.closeModal?.addEventListener('click', () => this.hideHelpModal());
        this.helpModal?.querySelector('.modal-backdrop')?.addEventListener('click', () => this.hideHelpModal());
        
        // Assistant popup
        this.closeAssistant?.addEventListener('click', () => this.hideAssistant());
        
        // Escape key to close modal/assistant
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                if (this.helpModal?.classList.contains('show')) {
                    this.hideHelpModal();
                }
                if (this.assistantPopup?.classList.contains('show')) {
                    this.hideAssistant();
                }
            }
        });
    }
    
    async createSession() {
        try {
            const response = await fetch('/api/session', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' }
            });
            const data = await response.json();
            
            if (data.success) {
                this.sessionId = data.session_id;
                console.log('Session created:', this.sessionId);
            }
        } catch (error) {
            console.error('Failed to create session:', error);
            this.sessionId = 'default';
        }
    }
    
    toggleListening() {
        if (this.isListening) {
            this.stopListening();
        } else {
            this.startListening();
        }
    }
    
    startListening() {
        if (!this.recognition) {
            this.showToast('Speech recognition not available', 'error');
            return;
        }
        
        try {
            this.recognition.start();
        } catch (error) {
            console.error('Failed to start recognition:', error);
        }
    }
    
    stopListening() {
        this.isListening = false;
        this.voiceBtn?.classList.remove('listening');
        this.updateStatus('Ready', 'ready');
        
        try {
            this.recognition?.stop();
        } catch (error) {
            // Ignore errors when stopping
        }
    }
    
    async processCommand(command) {
        command = command.trim();
        if (!command) return;
        
        // Debounce: prevent duplicate commands within debounce window
        const now = Date.now();
        const normalizedCommand = command.toLowerCase();
        
        if (normalizedCommand === this.lastCommand.toLowerCase() && 
            (now - this.lastCommandTime) < this.debounceDelay) {
            console.log('Debounced duplicate command:', command);
            return;
        }
        
        // Prevent concurrent processing
        if (this.isProcessing) {
            console.log('Already processing a command');
            return;
        }
        
        this.lastCommand = command;
        this.lastCommandTime = now;
        this.isProcessing = true;
        
        this.updateStatus('Processing...', 'processing');
        this.commandInput.value = '';
        
        // Show assistant popup with processing state
        this.showAssistantProcessing(command);
        
        try {
            const response = await fetch('/api/command', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    command: command,
                    session_id: this.sessionId,
                    voice_feedback: this.voiceFeedback?.checked || false
                })
            });
            
            const data = await response.json();
            
            // Add to history
            this.addToHistory(command, data.message, data.success);
            
            // Update assistant popup with result
            this.showAssistantResult(command, data);
            
            // Show toast notification
            this.showToast(data.message, data.success ? 'success' : 'error');
            
            // Update preview with cell highlighting
            if (data.data_preview) {
                this.updateTable(data.data_preview);
                
                // Highlight affected cells based on action
                this.highlightAffectedCells(data.action, data.params);
            }
            
            if (data.sheet_info) {
                this.updateSheetTabs(data.sheet_info);
            }
            
            // Render chart if a chart was created
            if (data.action && data.action.startsWith('chart_')) {
                this.renderChart(data.action, data.data_preview, data.params);
            }
            
            // Play audio feedback
            if (data.audio) {
                this.playAudio(data.audio);
            }
            
        } catch (error) {
            console.error('Error processing command:', error);
            this.showToast('Failed to process command. Please try again.', 'error');
            this.addToHistory(command, 'Error: Failed to process command', false);
            this.hideAssistant();
        }
        
        this.isProcessing = false;
        this.updateStatus('Ready', 'ready');
    }
    
    async uploadFile(file) {
        const formData = new FormData();
        formData.append('file', file);
        formData.append('session_id', this.sessionId);
        
        this.updateStatus('Uploading...', 'processing');
        
        try {
            const response = await fetch('/api/upload', {
                method: 'POST',
                body: formData
            });
            
            const data = await response.json();
            
            if (data.success) {
                this.showToast(data.message, 'success');
                
                if (data.data_preview) {
                    this.updateTable(data.data_preview);
                }
                
                if (data.sheet_info) {
                    this.updateSheetTabs(data.sheet_info);
                }
            } else {
                this.showToast(data.message, 'error');
            }
            
        } catch (error) {
            console.error('Error uploading file:', error);
            this.showToast('Failed to upload file', 'error');
        }
        
        this.updateStatus('Ready', 'ready');
    }
    
    async downloadFile() {
        try {
            const response = await fetch(`/api/download?session_id=${this.sessionId}`);
            
            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'workbook.xlsx';
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                a.remove();
                
                this.showToast('File downloaded successfully!', 'success');
            } else {
                this.showToast('Failed to download file', 'error');
            }
        } catch (error) {
            console.error('Error downloading file:', error);
            this.showToast('Failed to download file', 'error');
        }
    }
    
    async updatePreview() {
        try {
            const response = await fetch(`/api/preview?session_id=${this.sessionId}`);
            const data = await response.json();
            
            if (data.success) {
                if (data.data_preview) {
                    this.updateTable(data.data_preview);
                }
                
                if (data.sheet_info) {
                    this.updateSheetTabs(data.sheet_info);
                }
            }
        } catch (error) {
            console.error('Error fetching preview:', error);
        }
    }
    
    updateTable(data) {
        if (!data || data.length === 0) {
            this.tableHeader.innerHTML = '<th class="row-number">#</th><th>A</th><th>B</th><th>C</th><th>D</th><th>E</th>';
            this.tableBody.innerHTML = '<tr><td colspan="6" style="text-align: center; color: var(--text-muted);">No data yet. Try saying "Create sample sales data"</td></tr>';
            return;
        }
        
        // Determine number of columns
        const maxCols = Math.max(...data.map(row => row ? row.length : 0), 5);
        
        // Build header
        let headerHtml = '<th class="row-number">#</th>';
        for (let i = 0; i < maxCols; i++) {
            headerHtml += `<th>${String.fromCharCode(65 + i)}</th>`;
        }
        this.tableHeader.innerHTML = headerHtml;
        
        // Build body
        let bodyHtml = '';
        data.forEach((row, rowIndex) => {
            bodyHtml += `<tr><td class="row-number">${rowIndex + 1}</td>`;
            for (let i = 0; i < maxCols; i++) {
                const value = row && row[i] !== null && row[i] !== undefined ? row[i] : '';
                bodyHtml += `<td>${this.escapeHtml(String(value))}</td>`;
            }
            bodyHtml += '</tr>';
        });
        
        this.tableBody.innerHTML = bodyHtml || '<tr><td colspan="6" style="text-align: center; color: var(--text-muted);">No data</td></tr>';
    }
    
    updateSheetTabs(sheetInfo) {
        if (!sheetInfo || !sheetInfo.all_sheets) return;
        
        let tabsHtml = '';
        sheetInfo.all_sheets.forEach(sheet => {
            const isActive = sheet === sheetInfo.sheet_name ? 'active' : '';
            tabsHtml += `<button class="sheet-tab ${isActive}" data-sheet="${this.escapeHtml(sheet)}">${this.escapeHtml(sheet)}</button>`;
        });
        
        this.sheetTabs.innerHTML = tabsHtml;
        
        // Add click handlers to new tabs
        this.sheetTabs.querySelectorAll('.sheet-tab').forEach(tab => {
            tab.addEventListener('click', () => {
                const sheetName = tab.dataset.sheet;
                this.processCommand(`switch to sheet ${sheetName}`);
            });
        });
    }
    
    addToHistory(command, response, success) {
        const historyItem = { command, response, success, timestamp: new Date() };
        this.commandHistory.unshift(historyItem);
        
        // Keep only last 20 items
        if (this.commandHistory.length > 20) {
            this.commandHistory.pop();
        }
        
        this.renderHistory();
    }
    
    renderHistory() {
        if (this.commandHistory.length === 0) {
            this.historyList.innerHTML = `
                <div class="history-empty">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
                        <path d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"/>
                    </svg>
                    <p>No commands yet. Start by saying something!</p>
                </div>
            `;
            return;
        }
        
        let html = '';
        this.commandHistory.forEach(item => {
            const statusClass = item.success ? 'success' : 'error';
            // Truncate long responses
            const truncatedResponse = item.response.length > 100 
                ? item.response.substring(0, 100) + '...' 
                : item.response;
            
            html += `
                <div class="history-item ${statusClass}">
                    <div class="history-command">"${this.escapeHtml(item.command)}"</div>
                    <div class="history-response">${this.escapeHtml(truncatedResponse)}</div>
                </div>
            `;
        });
        
        this.historyList.innerHTML = html;
    }
    
    updateStatus(text, status) {
        if (!this.statusIndicator) return;
        
        this.statusIndicator.className = `status-indicator ${status}`;
        this.statusIndicator.querySelector('.status-text').textContent = text;
    }
    
    showToast(message, type = 'success') {
        if (!this.toast) return;
        
        this.toast.className = `toast ${type}`;
        this.toast.querySelector('.toast-message').textContent = message;
        this.toast.classList.add('show');
        
        // Auto hide after 4 seconds
        setTimeout(() => {
            this.toast.classList.remove('show');
        }, 4000);
    }
    
    showHelpModal() {
        this.helpModal?.classList.add('show');
    }
    
    hideHelpModal() {
        this.helpModal?.classList.remove('show');
    }
    
    playAudio(base64Audio) {
        try {
            const audio = new Audio(`data:audio/mp3;base64,${base64Audio}`);
            audio.play().catch(err => console.log('Audio playback failed:', err));
        } catch (error) {
            console.error('Error playing audio:', error);
        }
    }
    
    renderChart(chartType, dataPreview, params) {
        const chartContainer = document.getElementById('chartContainer');
        const chartCanvas = document.getElementById('chartCanvas');
        const chartTitle = document.getElementById('chartTitle');
        const closeChart = document.getElementById('closeChart');
        
        if (!chartContainer || !chartCanvas || !dataPreview || dataPreview.length < 2) {
            console.log('Cannot render chart: missing elements or insufficient data');
            return;
        }
        
        // Destroy existing chart if any
        if (this.currentChart) {
            this.currentChart.destroy();
        }
        
        // Extract labels and data from preview
        const labels = [];
        const values = [];
        
        // Skip header row (index 0), use column A for labels and column B for values
        for (let i = 1; i < dataPreview.length; i++) {
            const row = dataPreview[i];
            if (row && row[0] && row[1]) {
                labels.push(String(row[0]));
                const val = parseFloat(row[1]);
                if (!isNaN(val)) {
                    values.push(val);
                }
            }
        }
        
        if (labels.length === 0 || values.length === 0) {
            console.log('No valid data for chart');
            return;
        }
        
        // Chart colors
        const colors = [
            'rgba(0, 212, 170, 0.8)',
            'rgba(108, 92, 231, 0.8)',
            'rgba(255, 107, 107, 0.8)',
            'rgba(255, 193, 7, 0.8)',
            'rgba(0, 123, 255, 0.8)',
            'rgba(253, 126, 20, 0.8)',
            'rgba(111, 66, 193, 0.8)',
            'rgba(23, 162, 184, 0.8)'
        ];
        
        const borderColors = colors.map(c => c.replace('0.8', '1'));
        
        // Determine chart type
        let type = 'bar';
        if (chartType === 'chart_pie') {
            type = 'pie';
        } else if (chartType === 'chart_line') {
            type = 'line';
        }
        
        // Update title
        const title = params?.title || 'Data Visualization';
        chartTitle.textContent = title;
        
        // Create chart
        const ctx = chartCanvas.getContext('2d');
        this.currentChart = new Chart(ctx, {
            type: type,
            data: {
                labels: labels,
                datasets: [{
                    label: dataPreview[0] ? String(dataPreview[0][1] || 'Value') : 'Value',
                    data: values,
                    backgroundColor: type === 'pie' ? colors.slice(0, values.length) : colors[0],
                    borderColor: type === 'pie' ? borderColors.slice(0, values.length) : borderColors[0],
                    borderWidth: 2
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: {
                        display: type === 'pie',
                        position: 'right',
                        labels: {
                            color: '#e2e8f0'
                        }
                    },
                    title: {
                        display: false
                    }
                },
                scales: type !== 'pie' ? {
                    y: {
                        beginAtZero: true,
                        grid: {
                            color: 'rgba(255, 255, 255, 0.1)'
                        },
                        ticks: {
                            color: '#94a3b8'
                        }
                    },
                    x: {
                        grid: {
                            color: 'rgba(255, 255, 255, 0.1)'
                        },
                        ticks: {
                            color: '#94a3b8'
                        }
                    }
                } : {}
            }
        });
        
        // Show chart container
        chartContainer.style.display = 'block';
        
        // Scroll chart into view
        setTimeout(() => {
            chartContainer.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }, 100);
        
        // Add close button handler
        closeChart.onclick = () => {
            chartContainer.style.display = 'none';
            if (this.currentChart) {
                this.currentChart.destroy();
                this.currentChart = null;
            }
        };
        
        this.showToast(`${type.charAt(0).toUpperCase() + type.slice(1)} chart created!`, 'success');
    }
    
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
    
    // =============== ASSISTANT POPUP METHODS ===============
    
    showAssistantProcessing(command) {
        if (!this.assistantPopup) return;
        
        // Clear any existing timeout
        if (this.assistantTimeout) {
            clearTimeout(this.assistantTimeout);
        }
        
        this.processingIndicator.style.display = 'flex';
        this.assistantMessage.style.display = 'none';
        this.assistantAction.style.display = 'none';
        this.affectedCells.innerHTML = '';
        
        this.assistantStatus.textContent = 'Understanding...';
        this.assistantStatus.style.color = 'var(--warning)';
        
        this.assistantPopup.classList.add('show');
    }
    
    showAssistantResult(command, data) {
        if (!this.assistantPopup) return;
        
        this.processingIndicator.style.display = 'none';
        this.assistantMessage.style.display = 'block';
        this.assistantAction.style.display = 'flex';
        
        // Update status
        if (data.success) {
            this.assistantStatus.textContent = 'Action Completed';
            this.assistantStatus.style.color = 'var(--success)';
        } else {
            this.assistantStatus.textContent = 'Need Clarification';
            this.assistantStatus.style.color = 'var(--error)';
        }
        
        // Create context-aware message
        const contextMessage = this.createContextMessage(command, data);
        this.assistantMessage.innerHTML = contextMessage;
        
        // Update action text
        this.actionText.textContent = data.success ? data.message : 'Try rephrasing your command';
        
        // Show affected cells
        this.showAffectedCellTags(data.action, data.params);
        
        // Auto-hide after 5 seconds
        this.assistantTimeout = setTimeout(() => {
            this.hideAssistant();
        }, 5000);
    }
    
    createContextMessage(command, data) {
        const action = data.action || 'unknown';
        const params = data.params || {};
        
        let message = `I understood: "<strong>${this.escapeHtml(command)}</strong>"<br><br>`;
        
        // Add context-aware interpretation
        switch (action) {
            case 'write_cell':
                message += `📝 Writing "${params.value || ''}" to cell <strong>${params.cell || ''}</strong>`;
                break;
            case 'read_cell':
                message += `👁️ Reading value from cell <strong>${params.cell || ''}</strong>`;
                break;
            case 'format_bold':
            case 'format_italic':
            case 'format_color':
                message += `🎨 Applying formatting to <strong>${params.cell || params.start_cell || ''}</strong>`;
                break;
            case 'sum':
            case 'average':
            case 'max':
            case 'min':
                message += `🔢 Calculating ${action.toUpperCase()} for range <strong>${params.start || ''}:${params.end || ''}</strong>`;
                break;
            case 'chart_bar':
            case 'chart_pie':
            case 'chart_line':
                const chartType = action.replace('chart_', '');
                message += `📊 Creating ${chartType} chart from your data`;
                break;
            case 'sort':
                const order = params.ascending ? 'ascending' : 'descending';
                message += `↕️ Sorting data by column <strong>${params.column || 'A'}</strong> in ${order} order`;
                break;
            case 'filter':
                message += `🔍 Filtering data where column <strong>${params.column || ''}</strong> equals "${params.value || ''}"`;
                break;
            case 'create_sheet':
                message += `📄 Creating new sheet: <strong>${params.name || ''}</strong>`;
                break;
            case 'sample_data':
                message += `📋 Generating sample ${params.type || 'sales'} data for demonstration`;
                break;
            default:
                message += `Processing your request...`;
        }
        
        return message;
    }
    
    showAffectedCellTags(action, params) {
        if (!this.affectedCells) return;
        
        let cells = [];
        
        if (params.cell) cells.push(params.cell);
        if (params.start_cell) cells.push(params.start_cell);
        if (params.end_cell) cells.push(params.end_cell);
        if (params.start) cells.push(params.start);
        if (params.end) cells.push(params.end);
        if (params.result) cells.push(params.result);
        if (params.result_cell) cells.push(params.result_cell);
        
        if (cells.length === 0) {
            this.affectedCells.innerHTML = '';
            return;
        }
        
        const tagsHtml = cells.map(cell => 
            `<span class="affected-cell-tag">${this.escapeHtml(cell.toUpperCase())}</span>`
        ).join('');
        
        this.affectedCells.innerHTML = `<span style="font-size: 0.75rem; color: var(--text-muted); width: 100%;">Affected cells:</span>${tagsHtml}`;
    }
    
    hideAssistant() {
        if (this.assistantTimeout) {
            clearTimeout(this.assistantTimeout);
        }
        this.assistantPopup?.classList.remove('show');
    }
    
    // =============== CELL HIGHLIGHTING ===============
    
    highlightAffectedCells(action, params) {
        if (!params) return;
        
        console.log('Highlighting cells for action:', action, 'params:', params);
        
        // Determine highlight type based on action
        let highlightClass = 'highlight-affected';
        if (action === 'read_cell') highlightClass = 'highlight-read';
        if (['sum', 'sum_row', 'average', 'average_row', 'max', 'min', 'formula'].includes(action)) {
            highlightClass = 'highlight-formula';
        }
        
        // Small delay to ensure table is updated first
        setTimeout(() => {
            // For row operations
            if (params.row) {
                this.highlightRow(parseInt(params.row), highlightClass);
                return;
            }
            
            // Get cells to highlight
            const cellsToHighlight = [];
            
            if (params.cell) cellsToHighlight.push(params.cell);
            if (params.start) cellsToHighlight.push(params.start);
            if (params.end) cellsToHighlight.push(params.end);
            if (params.start_cell) cellsToHighlight.push(params.start_cell);
            if (params.end_cell) cellsToHighlight.push(params.end_cell);
            if (params.result) cellsToHighlight.push(params.result);
            if (params.result_cell) cellsToHighlight.push(params.result_cell);
            
            console.log('Cells to highlight:', cellsToHighlight);
            
            // Highlight cells in the table
            cellsToHighlight.forEach(cellRef => {
                this.highlightCell(cellRef, highlightClass);
            });
            
            // For range operations, highlight all cells in range
            if (params.start && params.end) {
                this.highlightRange(params.start, params.end, highlightClass);
            }
        }, 100);
    }
    
    highlightRow(rowNumber, highlightClass) {
        const tableRows = this.tableBody?.querySelectorAll('tr');
        if (!tableRows || rowNumber < 1 || rowNumber > tableRows.length) {
            console.log('Row not found:', rowNumber);
            return;
        }
        
        const cells = tableRows[rowNumber - 1]?.querySelectorAll('td');
        if (!cells) return;
        
        console.log('Highlighting row', rowNumber, 'with', cells.length, 'cells');
        
        cells.forEach((cell, index) => {
            if (index > 0) { // Skip row number column
                cell.classList.add(highlightClass);
                setTimeout(() => {
                    cell.classList.remove(highlightClass);
                }, 2000);
            }
        });
    }
    
    highlightCell(cellRef, highlightClass) {
        if (!cellRef) return;
        
        const match = cellRef.toString().toUpperCase().match(/([A-Z]+)(\d+)/);
        if (!match) {
            console.log('Invalid cell reference:', cellRef);
            return;
        }
        
        const col = match[1];
        const row = parseInt(match[2]);
        
        // Convert column letter to index (A=1, B=2, etc.)
        let colIndex = 0;
        for (let i = 0; i < col.length; i++) {
            colIndex = colIndex * 26 + (col.charCodeAt(i) - 64);
        }
        
        console.log(`Highlighting cell ${cellRef}: row=${row}, col=${colIndex}`);
        
        // Find the cell in the table
        const tableRows = this.tableBody?.querySelectorAll('tr');
        if (!tableRows || row < 1 || row > tableRows.length) {
            console.log('Row not found in table');
            return;
        }
        
        const cells = tableRows[row - 1]?.querySelectorAll('td');
        if (!cells || colIndex > cells.length) {
            console.log('Cell not found in row');
            return;
        }
        
        // colIndex is 1-based, but we need to account for row number column (index 0)
        const targetCell = cells[colIndex];
        if (targetCell) {
            console.log('Adding highlight class to cell');
            targetCell.classList.add(highlightClass);
            
            // Flash effect
            targetCell.style.transform = 'scale(1.05)';
            setTimeout(() => {
                targetCell.style.transform = '';
            }, 200);
            
            // Remove highlight after animation
            setTimeout(() => {
                targetCell.classList.remove(highlightClass);
            }, 2000);
        }
    }
    
    highlightRange(startRef, endRef, highlightClass) {
        const startMatch = startRef.toString().toUpperCase().match(/([A-Z]+)(\d+)/);
        const endMatch = endRef.toString().toUpperCase().match(/([A-Z]+)(\d+)/);
        
        if (!startMatch || !endMatch) return;
        
        const startCol = startMatch[1].charCodeAt(0) - 64;
        const endCol = endMatch[1].charCodeAt(0) - 64;
        const startRow = parseInt(startMatch[2]);
        const endRow = parseInt(endMatch[2]);
        
        console.log(`Highlighting range ${startRef}:${endRef}`);
        
        const tableRows = this.tableBody?.querySelectorAll('tr');
        if (!tableRows) return;
        
        for (let row = startRow; row <= endRow && row <= tableRows.length; row++) {
            const cells = tableRows[row - 1]?.querySelectorAll('td');
            if (!cells) continue;
            
            for (let col = startCol; col <= endCol && col < cells.length; col++) {
                const cell = cells[col];
                if (cell) {
                    cell.classList.add(highlightClass);
                    setTimeout(() => {
                        cell.classList.remove(highlightClass);
                    }, 2000);
                }
            }
        }
    }
}

// Initialize app when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    window.voiceExcelApp = new VoiceExcelApp();
});

