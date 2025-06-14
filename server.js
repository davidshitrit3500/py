require('dotenv').config();
const express = require('express');
const path = require('path');
const cors = require('cors');
const Imap = require('imap');
const { simpleParser } = require('mailparser');

const app = express();
app.use(cors());
app.use(express.json());

// Serve static files from current directory
app.use(express.static(__dirname));

// Home route
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Store IMAP connections
const connections = {};

// Connect to IMAP
app.post('/connect', async (req, res) => {
    const { email, password, server } = req.body;
    
    try {
        const imap = new Imap({
            user: email,
            password: password,
            host: server,
            port: 993,
            tls: true,
            tlsOptions: {
                rejectUnauthorized: false,
                servername: server
            },
            autotls: 'never'
        });

        // Connect to IMAP server
        await new Promise((resolve, reject) => {
            imap.once('ready', () => resolve());
            imap.once('error', (err) => reject(err));
            imap.connect();
        });

        // Store connection
        connections[email] = imap;

        res.json({ success: true, message: 'Connected successfully' });
    } catch (err) {
        console.error('IMAP connection error:', err);
        
        // Improved error handling
        let errorMessage = 'Connection failed';
        if (err.source === 'authentication') {
            errorMessage = 'Authentication failed: Invalid email or password';
        } else if (err.code === 'ECONNRESET') {
            errorMessage = 'Connection reset by server';
        } else if (err.message.includes('LOGIN')) {
            errorMessage = 'Login failed. Check your credentials.';
        }
        
        res.status(500).json({ 
            success: false, 
            message: errorMessage
        });
    }
});

// Get mailboxes
app.get('/mailboxes', async (req, res) => {
    const { email } = req.query;
    if (!email) {
        return res.status(400).json({ success: false, message: 'Email is required' });
    }

    const imap = connections[email];
    if (!imap) {
        return res.status(400).json({ success: false, message: 'Not connected' });
    }

    try {
        // Get mailboxes
        const boxes = await new Promise((resolve, reject) => {
            imap.getBoxes((err, boxes) => {
                if (err) reject(err);
                else resolve(boxes);
            });
        });

        // Process mailbox structure
        const processBoxes = (boxObj, prefix = '') => {
            const result = [];
            for (const [name, box] of Object.entries(boxObj)) {
                if (box.attribs.includes('\\Noselect')) continue;
                const fullPath = prefix + name;
                result.push({
                    name: name,
                    path: fullPath,
                    delimiter: box.delimiter,
                    attribs: box.attribs
                });

                if (box.children) {
                    result.push(...processBoxes(box.children, fullPath + box.delimiter));
                }
            }
            return result;
        };

        res.json({ success: true, mailboxes: processBoxes(boxes) });
    } catch (err) {
        console.error('Error loading mailboxes:', err);
        res.status(500).json({ 
            success: false, 
            message: 'Failed to load mailboxes'
        });
    }
});

// Get emails from a mailbox
app.post('/emails', async (req, res) => {
    const { email, mailbox } = req.body;
    if (!email || !mailbox) {
        return res.status(400).json({ success: false, message: 'Email and mailbox are required' });
    }

    const imap = connections[email];
    if (!imap) {
        return res.status(400).json({ success: false, message: 'Not connected' });
    }

    try {
        // Open mailbox
        await new Promise((resolve, reject) => {
            imap.openBox(mailbox, false, (err) => { // Changed to false to avoid auto-expunge
                if (err) reject(err);
                else resolve();
            });
        });

        // Fetch emails
        const emails = [];
        const fetch = imap.seq.fetch('1:5', { // Fetch first 5 emails
            bodies: ['HEADER.FIELDS (FROM SUBJECT DATE)', 'TEXT'],
            struct: true
        });

        fetch.on('message', (msg) => {
            let header = '';
            let text = '';

            msg.on('body', (stream, info) => {
                let buffer = '';
                stream.on('data', (chunk) => {
                    buffer += chunk.toString('utf8');
                });
                stream.on('end', () => {
                    if (info.which === 'HEADER.FIELDS (FROM SUBJECT DATE)') {
                        header = buffer;
                    } else if (info.which === 'TEXT') {
                        text = buffer;
                    }
                });
            });

            msg.once('end', () => {
                try {
                    // Simple parsing without mailparser
                    const fromMatch = /From: (.+)/i.exec(header);
                    const subjectMatch = /Subject: (.+)/i.exec(header);
                    const dateMatch = /Date: (.+)/i.exec(header);
                    
                    emails.push({
                        from: fromMatch ? fromMatch[1] : 'Unknown',
                        subject: subjectMatch ? subjectMatch[1] : '(No Subject)',
                        date: dateMatch ? dateMatch[1] : new Date().toISOString(),
                        preview: text.substring(0, 100) + '...'
                    });
                } catch (parseErr) {
                    console.error('Parse error:', parseErr);
                }
            });
        });

        fetch.once('error', (err) => {
            console.error('Fetch error:', err);
            res.status(500).json({ 
                success: false, 
                message: 'Error fetching emails'
            });
        });

        fetch.once('end', () => {
            res.json({ success: true, emails });
        });

    } catch (err) {
        console.error('Error in /emails:', err);
        res.status(500).json({ 
            success: false, 
            message: 'Error loading emails'
        });
    }
});

// Disconnect
app.post('/disconnect', (req, res) => {
    const { email } = req.body;
    if (!email) {
        return res.status(400).json({ success: false, message: 'Email is required' });
    }

    const imap = connections[email];
    if (imap) {
        imap.end();
        delete connections[email];
    }

    res.json({ success: true });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
    console.log(`Serving files from: ${__dirname}`);
});