document.addEventListener('DOMContentLoaded', function() {
    // DOM elements
    const emailInput = document.getElementById('email');
    const passwordInput = document.getElementById('password');
    const serverInput = document.getElementById('imap-server');
    const connectBtn = document.getElementById('connect-btn');
    const disconnectBtn = document.getElementById('disconnect-btn');
    const searchBtn = document.getElementById('search-btn');
    const refreshBtn = document.getElementById('refresh-btn');
    const mailboxSelect = document.getElementById('mailbox');
    const emailList = document.getElementById('email-list');
    const emailPreview = document.getElementById('email-preview');
    const statusIndicator = document.querySelector('.status-indicator');
    const connectionStatus = document.getElementById('connection-status');
    const userAvatar = document.getElementById('user-avatar');
    
    // Connection state
    let connected = false;
    let currentEmail = '';
    let currentMailbox = 'INBOX';
    let currentEmails = [];
    
    // Auto-detect IMAP server
    emailInput.addEventListener('blur', function() {
        if (this.value.includes('@')) {
            const domain = this.value.split('@')[1];
            const servers = {
                'gmail.com': 'imap.gmail.com',
                'yahoo.com': 'imap.mail.yahoo.com',
                'outlook.com': 'outlook.office365.com',
                'hotmail.com': 'outlook.office365.com',
                'icloud.com': 'imap.mail.me.com',
                'aol.com': 'imap.aol.com',
                't-online.de': 'imap.t-online.de'
            };
            serverInput.value = servers[domain] || `imap.${domain}`;
        }
    });
    
    // Connect to IMAP server
    connectBtn.addEventListener('click', async function() {
        const email = emailInput.value;
        const password = passwordInput.value;
        const server = serverInput.value;
        
        if (!email || !password || !server) {
            showError('Please enter email, password and IMAP server');
            return;
        }
        
        // Show loading state
        setLoading(true);
        showError('');
        
        try {
            // Send request to backend
            const response = await fetch('/connect', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email, password, server })
            });
            
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.message);
            }
            
            // Update UI
            connected = true;
            currentEmail = email;
            statusIndicator.classList.remove('offline');
            statusIndicator.innerHTML = '<i class="fas fa-circle"></i> Connected';
            statusIndicator.style.color = 'var(--success)';
            connectBtn.innerHTML = '<i class="fas fa-plug"></i> Disconnect';
            disconnectBtn.disabled = false;
            
            // Set user avatar
            const username = email.split('@')[0];
            userAvatar.textContent = username.charAt(0).toUpperCase();
            
            // Load mailboxes
            await loadMailboxes();
            
            // Load emails
            await loadEmails();
            
        } catch (error) {
            console.error('Connection error:', error);
            showError(`Connection failed: ${error.message}`);
        } finally {
            setLoading(false);
        }
    });
    
    // Load mailboxes
    async function loadMailboxes() {
        try {
            const response = await fetch(`/mailboxes?email=${encodeURIComponent(currentEmail)}`);
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.message);
            }
            
            // Update mailbox dropdown
            mailboxSelect.innerHTML = '';
            
            data.mailboxes.forEach(mailbox => {
                const option = document.createElement('option');
                option.value = mailbox.path;
                option.textContent = mailbox.path;
                mailboxSelect.appendChild(option);
            });
            
        } catch (error) {
            console.error('Mailbox error:', error);
            showError(`Error loading mailboxes: ${error.message}`);
        }
    }
    
    // Load emails
    async function loadEmails() {
        if (!connected) return;
        
        setLoading(true);
        const mailbox = mailboxSelect.value;
        
        try {
            // Send request to backend
            const response = await fetch('/emails', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email: currentEmail, mailbox })
            });
            
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.message);
            }
            
            // Update email list
            currentEmails = data.emails;
            renderEmailList(currentEmails);
            
        } catch (error) {
            console.error('Email error:', error);
            showError(`Error loading emails: ${error.message}`);
        } finally {
            setLoading(false);
        }
    }
    
    // Disconnect
    disconnectBtn.addEventListener('click', async function() {
        try {
            // Send disconnect request
            await fetch('/disconnect', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email: currentEmail })
            });
            
            connected = false;
            currentEmail = '';
            statusIndicator.classList.add('offline');
            statusIndicator.innerHTML = '<i class="fas fa-circle"></i> Disconnected';
            statusIndicator.style.color = 'var(--danger)';
            connectBtn.innerHTML = '<i class="fas fa-plug"></i> Connect';
            disconnectBtn.disabled = true;
            
            // Clear email list
            emailList.innerHTML = `
                <div class="empty-state">
                    <i class="fas fa-inbox"></i>
                    <h3>No Emails Loaded</h3>
                    <p>Connect to your IMAP account to view your emails</p>
                </div>
            `;
            
            // Clear preview
            emailPreview.innerHTML = `
                <div class="empty-state">
                    <i class="fas fa-envelope"></i>
                    <h3>No Email Selected</h3>
                    <p>Select an email from the list to view its content</p>
                </div>
            `;
            
        } catch (error) {
            console.error('Disconnect error:', error);
            showError(`Disconnect failed: ${error.message}`);
        }
    });
    
    // Render email list
    function renderEmailList(emails) {
        if (!emails || emails.length === 0) {
            emailList.innerHTML = `
                <div class="empty-state">
                    <i class="fas fa-inbox"></i>
                    <h3>No Emails Found</h3>
                    <p>No emails in this mailbox</p>
                </div>
            `;
            return;
        }
        
        emailList.innerHTML = '';
        
        emails.forEach((email, index) => {
            const sender = email.from;
            const subject = email.subject;
            const date = new Date(email.date).toLocaleDateString();
            const preview = email.preview;
            
            const emailItem = document.createElement('div');
            emailItem.className = 'email-item';
            if (index === 0) emailItem.classList.add('active');
            
            emailItem.innerHTML = `
                <div class="email-avatar">${sender.charAt(0).toUpperCase()}</div>
                <div class="email-content">
                    <div class="email-header">
                        <div class="email-sender">${sender}</div>
                        <div class="email-time">${date}</div>
                    </div>
                    <div class="email-subject">${subject}</div>
                    <div class="email-preview">${preview}</div>
                </div>
            `;
            
            emailItem.addEventListener('click', () => {
                document.querySelectorAll('.email-item').forEach(el => {
                    el.classList.remove('active');
                });
                emailItem.classList.add('active');
                renderEmailPreview(email);
            });
            
            emailList.appendChild(emailItem);
        });
        
        // Render preview for first email
        if (emails.length > 0) {
            renderEmailPreview(emails[0]);
        }
    }
    
    // Render email preview
    function renderEmailPreview(email) {
        const sender = email.from;
        const subject = email.subject;
        const date = new Date(email.date).toLocaleString();
        const body = email.preview;
        
        emailPreview.innerHTML = `
            <div class="email-preview-header">
                <div class="email-preview-subject">${subject}</div>
                <div class="email-preview-info">
                    <div class="email-preview-sender">
                        <div class="email-preview-avatar">${sender.charAt(0).toUpperCase()}</div>
                        <div>
                            <div>${sender}</div>
                            <div>to me</div>
                        </div>
                    </div>
                    <div>${date}</div>
                </div>
            </div>
            <div class="email-preview-body">
                <p>${body}</p>
            </div>
        `;
    }
    
    // Helper functions
    function setLoading(isLoading) {
        if (isLoading) {
            connectBtn.disabled = true;
            connectBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Connecting...';
        } else {
            connectBtn.disabled = false;
            connectBtn.innerHTML = '<i class="fas fa-plug"></i> Disconnect';
        }
    }
    
    function showError(message) {
        connectionStatus.textContent = message;
        connectionStatus.style.display = message ? 'block' : 'none';
    }
    
    // Auto-fill for testing
    emailInput.value = "guenesmerve@t-online.de";
    serverInput.value = "imap.t-online.de";
});