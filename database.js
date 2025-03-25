const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');

// Ensure the data directory exists
const dataDir = path.join(__dirname, 'data');
if (!fs.existsSync(dataDir)) {
    fs.mkdirSync(dataDir, { recursive: true });
}

// Create or open the database
const db = new sqlite3.Database(path.join(dataDir, 'email_app.db'));

// Initialize the database
function initializeDatabase() {
    return new Promise((resolve, reject) => {
        db.serialize(() => {
            // Create emails table
            db.run(`CREATE TABLE IF NOT EXISTS emails (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                recipient TEXT NOT NULL,
                subject TEXT NOT NULL,
                content TEXT NOT NULL,
                sent_date TEXT NOT NULL,
                contact_id TEXT,
                excel_file TEXT,
                sheet_name TEXT,
                line_number INTEGER,
                company TEXT,
                contact_name TEXT,
                excel_first_email_date TEXT,
                excel_first_email_status TEXT,
                excel_second_email_date TEXT,
                excel_second_email_status TEXT,
                excel_third_email_date TEXT,
                excel_third_email_status TEXT
            )`, (err) => {
                if (err) {
                    console.error('Error creating emails table:', err);
                    reject(err);
                    return;
                }
            });

            // Create follow-ups table
            db.run(`CREATE TABLE IF NOT EXISTS follow_ups (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email_id INTEGER NOT NULL,
                follow_up_date TEXT NOT NULL,
                status TEXT DEFAULT 'pending',
                follow_up_number INTEGER DEFAULT 1,
                completed_date TEXT,
                notes TEXT,
                FOREIGN KEY (email_id) REFERENCES emails (id)
            )`, (err) => {
                if (err) {
                    console.error('Error creating follow_ups table:', err);
                    reject(err);
                    return;
                }
                resolve();
            });
        });
    });
}

// Record a sent email
function recordEmail(emailData) {
    return new Promise((resolve, reject) => {
        const { recipient, subject, content, contactData = {} } = emailData;
        
        // Extract contact information
        const contactId = contactData.id || null;
        const excelFile = contactData.excelFile || null;
        const sheetName = contactData.sheetName || null;
        const lineNumber = contactData.lineNumber || null;
        const company = contactData.company || contactData[Object.keys(contactData)[0]] || null;
        const contactName = contactData.Name || contactData['שם'] || null;
        
        // Current date in ISO format
        const sentDate = new Date().toISOString();
        
        db.run(
            `INSERT INTO emails (recipient, subject, content, sent_date, contact_id, excel_file, sheet_name, line_number, company, contact_name) 
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
            [recipient, subject, content, sentDate, contactId, excelFile, sheetName, lineNumber, company, contactName],
            function(err) {
                if (err) {
                    console.error('Error recording email:', err);
                    reject(err);
                    return;
                }
                
                // Create follow-up reminder for 1 week later
                const emailId = this.lastID;
                const followUpDate = new Date();
                followUpDate.setDate(followUpDate.getDate() + 7); // 1 week later
                
                db.run(
                    `INSERT INTO follow_ups (email_id, follow_up_date, follow_up_number) 
                     VALUES (?, ?, ?)`,
                    [emailId, followUpDate.toISOString(), 1],
                    (err) => {
                        if (err) {
                            console.error('Error creating follow-up:', err);
                            reject(err);
                            return;
                        }
                        resolve(emailId);
                    }
                );
            }
        );
    });
}

// Get all sent emails
function getEmails() {
    return new Promise((resolve, reject) => {
        db.all(`SELECT * FROM emails ORDER BY sent_date DESC`, (err, rows) => {
            if (err) {
                console.error('Error getting sent emails:', err);
                reject(err);
                return;
            }
            resolve(rows);
        });
    });
}

// Get emails sent on a specific date
function getEmailsByDate(date) {
    return new Promise((resolve, reject) => {
        // Format date to match ISO string date part
        const dateStr = date.toISOString().split('T')[0];
        
        db.all(
            `SELECT * FROM emails WHERE sent_date LIKE ? ORDER BY sent_date DESC`,
            [`${dateStr}%`],
            (err, rows) => {
                if (err) {
                    console.error('Error getting emails by date:', err);
                    reject(err);
                    return;
                }
                resolve(rows);
            }
        );
    });
}

// Get emails sent to a specific recipient
function getEmailsByRecipient(recipient) {
    return new Promise((resolve, reject) => {
        db.all(
            `SELECT * FROM emails WHERE recipient LIKE ? ORDER BY sent_date DESC`,
            [`%${recipient}%`],
            (err, rows) => {
                if (err) {
                    console.error('Error getting emails by recipient:', err);
                    reject(err);
                    return;
                }
                resolve(rows);
            }
        );
    });
}

// Get all pending follow-ups
function getFollowUps() {
    return new Promise((resolve, reject) => {
        db.all(
            `SELECT f.*, e.recipient, e.subject, e.company, e.contact_name 
             FROM follow_ups f
             JOIN emails e ON f.email_id = e.id
             WHERE f.status = 'pending'
             ORDER BY f.follow_up_date ASC`,
            (err, rows) => {
                if (err) {
                    console.error('Error getting pending follow-ups:', err);
                    reject(err);
                    return;
                }
                resolve(rows);
            }
        );
    });
}

// Get follow-ups due on a specific date
function getFollowUpsByDate(date) {
    return new Promise((resolve, reject) => {
        // Format date to match ISO string date part
        const dateStr = date.toISOString().split('T')[0];
        
        db.all(
            `SELECT f.*, e.recipient, e.subject, e.company, e.contact_name 
             FROM follow_ups f
             JOIN emails e ON f.email_id = e.id
             WHERE f.follow_up_date LIKE ? AND f.status = 'pending'
             ORDER BY f.follow_up_date ASC`,
            [`${dateStr}%`],
            (err, rows) => {
                if (err) {
                    console.error('Error getting follow-ups by date:', err);
                    reject(err);
                    return;
                }
                resolve(rows);
            }
        );
    });
}

// Mark a follow-up as completed and create next follow-up if needed
function completeFollowUp(followUpId, notes = '') {
    return new Promise((resolve, reject) => {
        // First, get the follow-up details
        db.get(`SELECT * FROM follow_ups WHERE id = ?`, [followUpId], (err, followUp) => {
            if (err) {
                console.error('Error getting follow-up details:', err);
                reject(err);
                return;
            }
            
            if (!followUp) {
                reject(new Error('Follow-up not found'));
                return;
            }
            
            // Update the current follow-up as completed
            const completedDate = new Date().toISOString();
            
            db.run(
                `UPDATE follow_ups SET status = 'completed', completed_date = ?, notes = ? WHERE id = ?`,
                [completedDate, notes, followUpId],
                function(err) {
                    if (err) {
                        console.error('Error updating follow-up:', err);
                        reject(err);
                        return;
                    }
                    
                    // If this was not the 3rd follow-up, create a new one
                    if (followUp.follow_up_number < 3) {
                        const nextFollowUpNumber = followUp.follow_up_number + 1;
                        const nextFollowUpDate = new Date();
                        nextFollowUpDate.setDate(nextFollowUpDate.getDate() + 7); // 1 week later
                        
                        db.run(
                            `INSERT INTO follow_ups (email_id, follow_up_date, follow_up_number) 
                             VALUES (?, ?, ?)`,
                            [followUp.email_id, nextFollowUpDate.toISOString(), nextFollowUpNumber],
                            (err) => {
                                if (err) {
                                    console.error('Error creating next follow-up:', err);
                                    reject(err);
                                    return;
                                }
                                resolve({ completed: true, nextFollowUp: true });
                            }
                        );
                    } else {
                        // This was the 3rd follow-up, no more needed
                        resolve({ completed: true, nextFollowUp: false });
                    }
                }
            );
        });
    });
}

// Get calendar data for a month
function getCalendarData(year, month) {
    return new Promise((resolve, reject) => {
        // Month is 0-indexed in JavaScript Date but we want to accept 1-indexed
        const startDate = new Date(year, month - 1, 1);
        const endDate = new Date(year, month, 0); // Last day of the month
        
        const startDateStr = startDate.toISOString().split('T')[0];
        const endDateStr = endDate.toISOString().split('T')[0];
        
        // Get emails sent in the month
        const emailsPromise = new Promise((resolve, reject) => {
            db.all(
                `SELECT id, recipient, subject, sent_date, company, contact_name 
                 FROM emails 
                 WHERE sent_date >= ? AND sent_date <= ?`,
                [`${startDateStr}T00:00:00.000Z`, `${endDateStr}T23:59:59.999Z`],
                (err, rows) => {
                    if (err) {
                        console.error('Error getting emails for calendar:', err);
                        reject(err);
                        return;
                    }
                    resolve(rows);
                }
            );
        });
        
        // Get follow-ups due in the month
        const followUpsPromise = new Promise((resolve, reject) => {
            db.all(
                `SELECT f.id, f.follow_up_date, f.status, f.follow_up_number, 
                        e.recipient, e.subject, e.company, e.contact_name 
                 FROM follow_ups f
                 JOIN emails e ON f.email_id = e.id
                 WHERE f.follow_up_date >= ? AND f.follow_up_date <= ?`,
                [`${startDateStr}T00:00:00.000Z`, `${endDateStr}T23:59:59.999Z`],
                (err, rows) => {
                    if (err) {
                        console.error('Error getting follow-ups for calendar:', err);
                        reject(err);
                        return;
                    }
                    resolve(rows);
                }
            );
        });
        
        // Combine the results
        Promise.all([emailsPromise, followUpsPromise])
            .then(([emails, followUps]) => {
                // Format the data by day
                const calendarData = {};
                
                // Process emails
                emails.forEach(email => {
                    const date = new Date(email.sent_date);
                    const day = date.getDate();
                    
                    if (!calendarData[day]) {
                        calendarData[day] = { emails: [], followUps: [] };
                    }
                    
                    calendarData[day].emails.push(email);
                });
                
                // Process follow-ups
                followUps.forEach(followUp => {
                    const date = new Date(followUp.follow_up_date);
                    const day = date.getDate();
                    
                    if (!calendarData[day]) {
                        calendarData[day] = { emails: [], followUps: [] };
                    }
                    
                    calendarData[day].followUps.push(followUp);
                });
                
                resolve({ success: true, calendarData });
            })
            .catch(err => {
                console.error('Error processing calendar data:', err);
                reject(err);
            });
    });
}

// Get all unique companies from contacts
function getAllCompanies() {
    return new Promise((resolve, reject) => {
        db.all(`SELECT DISTINCT company FROM emails WHERE company IS NOT NULL AND company != ''`, [], (err, rows) => {
            if (err) {
                console.error('Error fetching companies:', err);
                reject(err);
                return;
            }
            
            // Extract company names and sort them
            const companies = rows.map(row => row.company).sort();
            resolve(companies);
        });
    });
}

// Get calendar data for a specific company
function getCompanyCalendarData(year, month, company) {
    return new Promise((resolve, reject) => {
        // Month is 0-indexed in JavaScript Date but we want to accept 1-indexed
        const startDate = new Date(year, month - 1, 1);
        const endDate = new Date(year, month, 0); // Last day of the month
        
        const startDateStr = startDate.toISOString().split('T')[0];
        const endDateStr = endDate.toISOString().split('T')[0];
        
        // Get emails sent in the month for the specific company
        const emailsPromise = new Promise((resolve, reject) => {
            db.all(
                `SELECT id, recipient, subject, sent_date, company, contact_name,
                        excel_first_email_date, excel_first_email_status,
                        excel_second_email_date, excel_second_email_status,
                        excel_third_email_date, excel_third_email_status
                 FROM emails 
                 WHERE (sent_date >= ? AND sent_date <= ?) AND company = ?`,
                [`${startDateStr}T00:00:00.000Z`, `${endDateStr}T23:59:59.999Z`, company],
                (err, rows) => {
                    if (err) {
                        console.error('Error getting emails for company calendar:', err);
                        reject(err);
                        return;
                    }
                    resolve(rows);
                }
            );
        });
        
        // Get follow-ups due in the month for the specific company
        const followUpsPromise = new Promise((resolve, reject) => {
            db.all(
                `SELECT f.id, f.follow_up_date, f.status, f.follow_up_number, 
                        e.recipient, e.subject, e.company, e.contact_name,
                        e.excel_first_email_date, e.excel_first_email_status,
                        e.excel_second_email_date, e.excel_second_email_status,
                        e.excel_third_email_date, e.excel_third_email_status
                 FROM follow_ups f
                 JOIN emails e ON f.email_id = e.id
                 WHERE (f.follow_up_date >= ? AND f.follow_up_date <= ?) AND e.company = ?`,
                [`${startDateStr}T00:00:00.000Z`, `${endDateStr}T23:59:59.999Z`, company],
                (err, rows) => {
                    if (err) {
                        console.error('Error getting follow-ups for company calendar:', err);
                        reject(err);
                        return;
                    }
                    resolve(rows);
                }
            );
        });
        
        // Get Excel-based follow-ups for the company
        const excelFollowUpsPromise = new Promise((resolve, reject) => {
            db.all(
                `SELECT id, recipient, subject, company, contact_name,
                        excel_first_email_date, excel_first_email_status,
                        excel_second_email_date, excel_second_email_status,
                        excel_third_email_date, excel_third_email_status
                 FROM emails
                 WHERE company = ? AND (
                    (excel_first_email_date >= ? AND excel_first_email_date <= ?) OR
                    (excel_second_email_date >= ? AND excel_second_email_date <= ?) OR
                    (excel_third_email_date >= ? AND excel_third_email_date <= ?)
                 )`,
                [
                    company,
                    `${startDateStr}T00:00:00.000Z`, `${endDateStr}T23:59:59.999Z`,
                    `${startDateStr}T00:00:00.000Z`, `${endDateStr}T23:59:59.999Z`,
                    `${startDateStr}T00:00:00.000Z`, `${endDateStr}T23:59:59.999Z`
                ],
                (err, rows) => {
                    if (err) {
                        console.error('Error getting Excel follow-ups for company calendar:', err);
                        reject(err);
                        return;
                    }
                    resolve(rows);
                }
            );
        });
        
        // Combine the results
        Promise.all([emailsPromise, followUpsPromise, excelFollowUpsPromise])
            .then(([emails, followUps, excelFollowUps]) => {
                // Format the data by day
                const calendarData = {};
                
                // Process emails
                emails.forEach(email => {
                    const date = new Date(email.sent_date);
                    const day = date.getDate();
                    
                    if (!calendarData[day]) {
                        calendarData[day] = { emails: [], followUps: [], excelFollowUps: [] };
                    }
                    
                    calendarData[day].emails.push(email);
                });
                
                // Process follow-ups
                followUps.forEach(followUp => {
                    const date = new Date(followUp.follow_up_date);
                    const day = date.getDate();
                    
                    if (!calendarData[day]) {
                        calendarData[day] = { emails: [], followUps: [], excelFollowUps: [] };
                    }
                    
                    calendarData[day].followUps.push(followUp);
                });
                
                // Process Excel follow-ups
                excelFollowUps.forEach(excelData => {
                    // Check first email
                    if (excelData.excel_first_email_date) {
                        const date = new Date(excelData.excel_first_email_date);
                        if (date.getMonth() + 1 === parseInt(month) && date.getFullYear() === parseInt(year)) {
                            const day = date.getDate();
                            
                            if (!calendarData[day]) {
                                calendarData[day] = { emails: [], followUps: [], excelFollowUps: [] };
                            }
                            
                            calendarData[day].excelFollowUps.push({
                                ...excelData,
                                followUpNumber: 1,
                                followUpDate: excelData.excel_first_email_date,
                                status: excelData.excel_first_email_status
                            });
                        }
                    }
                    
                    // Check second email
                    if (excelData.excel_second_email_date) {
                        const date = new Date(excelData.excel_second_email_date);
                        if (date.getMonth() + 1 === parseInt(month) && date.getFullYear() === parseInt(year)) {
                            const day = date.getDate();
                            
                            if (!calendarData[day]) {
                                calendarData[day] = { emails: [], followUps: [], excelFollowUps: [] };
                            }
                            
                            calendarData[day].excelFollowUps.push({
                                ...excelData,
                                followUpNumber: 2,
                                followUpDate: excelData.excel_second_email_date,
                                status: excelData.excel_second_email_status
                            });
                        }
                    }
                    
                    // Check third email
                    if (excelData.excel_third_email_date) {
                        const date = new Date(excelData.excel_third_email_date);
                        if (date.getMonth() + 1 === parseInt(month) && date.getFullYear() === parseInt(year)) {
                            const day = date.getDate();
                            
                            if (!calendarData[day]) {
                                calendarData[day] = { emails: [], followUps: [], excelFollowUps: [] };
                            }
                            
                            calendarData[day].excelFollowUps.push({
                                ...excelData,
                                followUpNumber: 3,
                                followUpDate: excelData.excel_third_email_date,
                                status: excelData.excel_third_email_status
                            });
                        }
                    }
                });
                
                resolve({ success: true, calendarData });
            })
            .catch(err => {
                console.error('Error processing company calendar data:', err);
                reject(err);
            });
    });
}

// Get calendar data for a specific company
function getCalendarDataForCompany(year, month, company) {
    return new Promise((resolve, reject) => {
        // Get the start and end dates for the month
        const startDate = new Date(year, month - 1, 1);
        const endDate = new Date(year, month, 0);
        
        const startDateStr = startDate.toISOString().split('T')[0];
        const endDateStr = endDate.toISOString().split('T')[0];
        
        // Create an empty calendar object with days of the month as keys
        const calendarData = {};
        for (let i = 1; i <= endDate.getDate(); i++) {
            calendarData[i] = {
                emails: [],
                followUps: [],
                excelFollowUps: []
            };
        }
        
        // Get emails for the company in the specified month
        const emailsPromise = new Promise((resolve, reject) => {
            db.all(
                `SELECT e.* FROM emails e
                 JOIN contacts c ON e.recipient = c.email
                 WHERE c.company = ? AND date(e.sent_date) BETWEEN date(?) AND date(?)`,
                [company, startDateStr, endDateStr],
                (err, rows) => {
                    if (err) {
                        console.error('Error fetching emails for company calendar:', err);
                        reject(err);
                        return;
                    }
                    
                    // Add emails to the calendar data
                    rows.forEach(email => {
                        const emailDate = new Date(email.sent_date);
                        const day = emailDate.getDate();
                        
                        if (calendarData[day]) {
                            calendarData[day].emails.push(email);
                        }
                    });
                    
                    resolve();
                }
            );
        });
        
        // Get follow-ups for the company in the specified month
        const followUpsPromise = new Promise((resolve, reject) => {
            db.all(
                `SELECT f.* FROM follow_ups f
                 JOIN contacts c ON f.recipient = c.email
                 WHERE c.company = ? AND date(f.follow_up_date) BETWEEN date(?) AND date(?)`,
                [company, startDateStr, endDateStr],
                (err, rows) => {
                    if (err) {
                        console.error('Error fetching follow-ups for company calendar:', err);
                        reject(err);
                        return;
                    }
                    
                    // Add follow-ups to the calendar data
                    rows.forEach(followUp => {
                        const followUpDate = new Date(followUp.follow_up_date);
                        const day = followUpDate.getDate();
                        
                        if (calendarData[day]) {
                            calendarData[day].followUps.push(followUp);
                        }
                    });
                    
                    resolve();
                }
            );
        });
        
        // Get Excel follow-ups for the company in the specified month
        const excelFollowUpsPromise = new Promise((resolve, reject) => {
            // Query for first email follow-ups
            const firstEmailPromise = new Promise((resolve, reject) => {
                db.all(
                    `SELECT e.*, c.company, c.name as contact_name 
                     FROM emails e
                     JOIN contacts c ON e.recipient = c.email
                     WHERE c.company = ? 
                     AND e.excel_first_email_date IS NOT NULL
                     AND date(e.excel_first_email_date) BETWEEN date(?) AND date(?)`,
                    [company, startDateStr, endDateStr],
                    (err, rows) => {
                        if (err) {
                            console.error('Error fetching first Excel follow-ups:', err);
                            reject(err);
                            return;
                        }
                        
                        // Add Excel follow-ups to the calendar data
                        rows.forEach(email => {
                            const followUpDate = new Date(email.excel_first_email_date);
                            const day = followUpDate.getDate();
                            
                            if (calendarData[day]) {
                                calendarData[day].excelFollowUps.push({
                                    id: email.id,
                                    contact_name: email.contact_name,
                                    company: email.company,
                                    recipient: email.recipient,
                                    followUpDate: email.excel_first_email_date,
                                    status: email.excel_first_email_status,
                                    followUpNumber: 1,
                                    excel_file: email.excel_file
                                });
                            }
                        });
                        
                        resolve();
                    }
                );
            });
            
            // Query for second email follow-ups
            const secondEmailPromise = new Promise((resolve, reject) => {
                db.all(
                    `SELECT e.*, c.company, c.name as contact_name 
                     FROM emails e
                     JOIN contacts c ON e.recipient = c.email
                     WHERE c.company = ? 
                     AND e.excel_second_email_date IS NOT NULL
                     AND date(e.excel_second_email_date) BETWEEN date(?) AND date(?)`,
                    [company, startDateStr, endDateStr],
                    (err, rows) => {
                        if (err) {
                            console.error('Error fetching second Excel follow-ups:', err);
                            reject(err);
                            return;
                        }
                        
                        // Add Excel follow-ups to the calendar data
                        rows.forEach(email => {
                            const followUpDate = new Date(email.excel_second_email_date);
                            const day = followUpDate.getDate();
                            
                            if (calendarData[day]) {
                                calendarData[day].excelFollowUps.push({
                                    id: email.id,
                                    contact_name: email.contact_name,
                                    company: email.company,
                                    recipient: email.recipient,
                                    followUpDate: email.excel_second_email_date,
                                    status: email.excel_second_email_status,
                                    followUpNumber: 2,
                                    excel_file: email.excel_file
                                });
                            }
                        });
                        
                        resolve();
                    }
                );
            });
            
            // Query for third email follow-ups
            const thirdEmailPromise = new Promise((resolve, reject) => {
                db.all(
                    `SELECT e.*, c.company, c.name as contact_name 
                     FROM emails e
                     JOIN contacts c ON e.recipient = c.email
                     WHERE c.company = ? 
                     AND e.excel_third_email_date IS NOT NULL
                     AND date(e.excel_third_email_date) BETWEEN date(?) AND date(?)`,
                    [company, startDateStr, endDateStr],
                    (err, rows) => {
                        if (err) {
                            console.error('Error fetching third Excel follow-ups:', err);
                            reject(err);
                            return;
                        }
                        
                        // Add Excel follow-ups to the calendar data
                        rows.forEach(email => {
                            const followUpDate = new Date(email.excel_third_email_date);
                            const day = followUpDate.getDate();
                            
                            if (calendarData[day]) {
                                calendarData[day].excelFollowUps.push({
                                    id: email.id,
                                    contact_name: email.contact_name,
                                    company: email.company,
                                    recipient: email.recipient,
                                    followUpDate: email.excel_third_email_date,
                                    status: email.excel_third_email_status,
                                    followUpNumber: 3,
                                    excel_file: email.excel_file
                                });
                            }
                        });
                        
                        resolve();
                    }
                );
            });
            
            // Resolve when all Excel follow-up queries are complete
            Promise.all([firstEmailPromise, secondEmailPromise, thirdEmailPromise])
                .then(() => resolve())
                .catch(err => reject(err));
        });
        
        // Resolve when all queries are complete
        Promise.all([emailsPromise, followUpsPromise, excelFollowUpsPromise])
            .then(() => resolve(calendarData))
            .catch(err => reject(err));
    });
}

// Import follow-up data from Excel
function importFollowUpDataFromExcel(excelFile, sheetName, contactData) {
    return new Promise((resolve, reject) => {
        try {
            // Check if we have follow-up data in the Excel
            if (!contactData) {
                resolve(null);
                return;
            }
            
            // Extract follow-up data from Excel columns
            // Columns M and N for first email, O and P for second, Q and R for third
            const firstEmailDate = contactData['M'] || null;
            const firstEmailStatus = contactData['N'] || null;
            const secondEmailDate = contactData['O'] || null;
            const secondEmailStatus = contactData['P'] || null;
            const thirdEmailDate = contactData['Q'] || null;
            const thirdEmailStatus = contactData['R'] || null;
            
            // If no follow-up data, return
            if (!firstEmailDate && !secondEmailDate && !thirdEmailDate) {
                resolve(null);
                return;
            }
            
            // Format the data
            const followUpData = {
                excelFile,
                sheetName,
                lineNumber: contactData.lineNumber,
                company: contactData.company || contactData[Object.keys(contactData)[0]] || null,
                contactName: contactData.Name || contactData['שם'] || null,
                firstEmail: {
                    date: firstEmailDate ? new Date(firstEmailDate) : null,
                    status: firstEmailStatus
                },
                secondEmail: {
                    date: secondEmailDate ? new Date(secondEmailDate) : null,
                    status: secondEmailStatus
                },
                thirdEmail: {
                    date: thirdEmailDate ? new Date(thirdEmailDate) : null,
                    status: thirdEmailStatus
                }
            };
            
            resolve(followUpData);
        } catch (error) {
            console.error('Error importing follow-up data from Excel:', error);
            reject(error);
        }
    });
}

// Update email with Excel follow-up data
function updateEmailWithExcelData(emailId, followUpData) {
    return new Promise((resolve, reject) => {
        if (!followUpData) {
            resolve();
            return;
        }
        
        db.run(
            `UPDATE emails SET 
                excel_first_email_date = ?,
                excel_first_email_status = ?,
                excel_second_email_date = ?,
                excel_second_email_status = ?,
                excel_third_email_date = ?,
                excel_third_email_status = ?
             WHERE id = ?`,
            [
                followUpData.firstEmail.date ? followUpData.firstEmail.date.toISOString() : null,
                followUpData.firstEmail.status,
                followUpData.secondEmail.date ? followUpData.secondEmail.date.toISOString() : null,
                followUpData.secondEmail.status,
                followUpData.thirdEmail.date ? followUpData.thirdEmail.date.toISOString() : null,
                followUpData.thirdEmail.status,
                emailId
            ],
            function(err) {
                if (err) {
                    console.error('Error updating email with Excel data:', err);
                    reject(err);
                    return;
                }
                resolve();
            }
        );
    });
}

module.exports = {
    initializeDatabase,
    recordEmail,
    getEmails,
    getEmailsByDate,
    getEmailsByRecipient,
    getFollowUps,
    getFollowUpsByDate,
    completeFollowUp,
    getCalendarData,
    importFollowUpDataFromExcel,
    updateEmailWithExcelData,
    getCompanyCalendarData,
    getAllCompanies,
    getCalendarDataForCompany
};
