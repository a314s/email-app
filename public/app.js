// DOM Elements
const elements = {
    // Google Sheet Configuration
    spreadsheetId: document.getElementById('spreadsheetId'),
    loadSheets: document.getElementById('loadSheets'),
    sheetSelect: document.getElementById('sheetSelect'),
    loadContacts: document.getElementById('loadContacts'),
    
    // Email Templates
    templateFile: document.getElementById('templateFile'),
    uploadTemplate: document.getElementById('uploadTemplate'),
    templatesList: document.getElementById('templatesList'),
    
    // Email Configuration
    fromEmail: document.getElementById('fromEmail'),
    emailMethod: document.getElementById('emailMethod'),
    
    // Search
    lineNumber: document.getElementById('lineNumber'),
    searchByLine: document.getElementById('searchByLine'),
    contactName: document.getElementById('contactName'),
    searchByName: document.getElementById('searchByName'),
    
    // Contact Info
    contactInfoSection: document.getElementById('contactInfoSection'),
    contactDetails: document.getElementById('contactDetails'),
    contactNameDisplay: document.getElementById('contactName'),
    contactCompany: document.getElementById('contactCompany'),
    lineInfo: document.getElementById('lineInfo'),
    contactDetailsGrid: document.getElementById('contactDetailsGrid'),
    generateEmail: document.getElementById('generateEmail'),
    viewAllData: document.getElementById('viewAllData'),
    noContactFound: document.getElementById('noContactFound'),
    
    // Email Preview
    emailPreviewSection: document.getElementById('emailPreviewSection'),
    emailTo: document.getElementById('emailTo'),
    emailSubject: document.getElementById('emailSubject'),
    emailContent: document.getElementById('emailContent'),
    sendEmail: document.getElementById('sendEmail'),
    openInOutlook: document.getElementById('openInOutlook'),
    backToContact: document.getElementById('backToContact'),
    
    // Data Table
    dataTableSection: document.getElementById('dataTableSection'),
    contactsTable: document.getElementById('contactsTable'),
    backFromTable: document.getElementById('backFromTable'),
    
    // UI Elements
    toast: document.getElementById('toast'),
    loadingOverlay: document.getElementById('loadingOverlay')
};

// App State
const state = {
    spreadsheetId: '',
    selectedSheet: '',
    sheetData: null,
    currentContact: null,
    templates: [],
    selectedTemplate: '',
    emailContent: ''
};

// API Functions
const api = {
    async fetchSheets(spreadsheetId) {
        showLoading(true);
        try {
            const response = await fetch(`/api/sheets/${spreadsheetId}`);
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to fetch sheets');
            }
            
            return data.sheets;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async fetchSheetData(spreadsheetId, sheetName) {
        showLoading(true);
        try {
            const response = await fetch(`/api/sheets/${spreadsheetId}/${sheetName}`);
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to fetch sheet data');
            }
            
            return data;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async fetchContactByLine(spreadsheetId, sheetName, lineNumber) {
        showLoading(true);
        try {
            const response = await fetch(`/api/contact/line/${spreadsheetId}/${sheetName}/${lineNumber}`);
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Contact not found');
            }
            
            return data.contact;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async fetchContactByName(spreadsheetId, sheetName, name) {
        showLoading(true);
        try {
            const response = await fetch(`/api/contact/name/${spreadsheetId}/${sheetName}/${encodeURIComponent(name)}`);
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Contact not found');
            }
            
            return data.contact;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async uploadTemplate(file) {
        showLoading(true);
        try {
            const formData = new FormData();
            formData.append('template', file);
            
            const response = await fetch('/api/template/upload', {
                method: 'POST',
                body: formData
            });
            
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to upload template');
            }
            
            return data;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async fetchTemplates() {
        showLoading(true);
        try {
            const response = await fetch('/api/templates');
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to fetch templates');
            }
            
            return data.templates;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async generateEmail(templateName, contactData) {
        showLoading(true);
        try {
            const response = await fetch('/api/generate-email', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ templateName, contactData })
            });
            
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to generate email');
            }
            
            return data;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async sendEmail(to, subject, content, from) {
        showLoading(true);
        try {
            const response = await fetch('/api/send-email', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ to, subject, content, from })
            });
            
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to send email');
            }
            
            return data;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async getOutlookUrl(to, subject, body) {
        showLoading(true);
        try {
            const response = await fetch('/api/open-outlook', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ to, subject, body })
            });
            
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to generate Outlook URL');
            }
            
            return data.mailtoUrl;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    }
};

// UI Helper Functions
function showLoading(show) {
    elements.loadingOverlay.classList.toggle('hidden', !show);
}

function showToast(message, type = 'info') {
    elements.toast.textContent = message;
    elements.toast.className = 'toast';
    elements.toast.classList.add(type);
    elements.toast.classList.remove('hidden');
    
    setTimeout(() => {
        elements.toast.classList.add('hidden');
    }, 3000);
}

function showSection(section) {
    // Hide all sections
    elements.contactInfoSection.classList.add('hidden');
    elements.emailPreviewSection.classList.add('hidden');
    elements.dataTableSection.classList.add('hidden');
    
    // Show the requested section
    section.classList.remove('hidden');
}

function populateContactDetails(contact) {
    // Clear previous data
    elements.contactDetailsGrid.innerHTML = '';
    
    // Set contact header info
    elements.contactNameDisplay.textContent = contact.Name || 'Unknown Contact';
    elements.contactCompany.textContent = contact.Company || '';
    elements.lineInfo.textContent = contact.lineNumber ? `Line: ${contact.lineNumber}` : '';
    
    // Create detail items for each property
    for (const [key, value] of Object.entries(contact)) {
        if (key === 'lineNumber') continue; // Skip line number as it's shown in the header
        
        const detailItem = document.createElement('div');
        detailItem.className = 'detail-item';
        
        const label = document.createElement('div');
        label.className = 'label';
        label.textContent = key;
        
        const valueElement = document.createElement('div');
        valueElement.className = 'value';
        valueElement.textContent = value || 'N/A';
        
        detailItem.appendChild(label);
        detailItem.appendChild(valueElement);
        elements.contactDetailsGrid.appendChild(detailItem);
    }
    
    // Show contact details and hide "not found" message
    elements.contactDetails.classList.remove('hidden');
    elements.noContactFound.classList.add('hidden');
    
    // Enable generate email button if templates are available
    elements.generateEmail.disabled = state.templates.length === 0;
}

function showContactNotFound() {
    elements.contactDetails.classList.add('hidden');
    elements.noContactFound.classList.remove('hidden');
}

function populateSheetSelect(sheets) {
    // Clear previous options
    elements.sheetSelect.innerHTML = '';
    
    // Add default option
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = 'Select a sheet...';
    elements.sheetSelect.appendChild(defaultOption);
    
    // Add options for each sheet
    sheets.forEach(sheet => {
        const option = document.createElement('option');
        option.value = sheet;
        option.textContent = sheet;
        elements.sheetSelect.appendChild(option);
    });
    
    // Enable the select
    elements.sheetSelect.disabled = false;
}

function populateTemplatesList(templates) {
    // Clear previous templates
    elements.templatesList.innerHTML = '';
    
    if (templates.length === 0) {
        const li = document.createElement('li');
        li.textContent = 'No templates available';
        elements.templatesList.appendChild(li);
        return;
    }
    
    // Add each template to the list
    templates.forEach(template => {
        const li = document.createElement('li');
        
        const templateName = document.createElement('span');
        templateName.textContent = template.filename;
        
        const selectButton = document.createElement('button');
        selectButton.innerHTML = '<i class="fas fa-check"></i> Select';
        selectButton.addEventListener('click', () => {
            state.selectedTemplate = template.filename;
            
            // Update all template list items to show which is selected
            document.querySelectorAll('#templatesList li').forEach(item => {
                item.classList.remove('selected');
            });
            li.classList.add('selected');
            
            showToast(`Selected template: ${template.filename}`, 'success');
            
            // Enable generate email button if a contact is selected
            if (state.currentContact) {
                elements.generateEmail.disabled = false;
            }
        });
        
        li.appendChild(templateName);
        li.appendChild(selectButton);
        elements.templatesList.appendChild(li);
    });
}

function populateDataTable(data) {
    const { headers, data: groupedData } = data;
    
    // Clear previous table
    elements.contactsTable.innerHTML = '';
    
    // Create table header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    
    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    
    // Add action column
    const actionHeader = document.createElement('th');
    actionHeader.textContent = 'Actions';
    headerRow.appendChild(actionHeader);
    
    thead.appendChild(headerRow);
    elements.contactsTable.appendChild(thead);
    
    // Create table body
    const tbody = document.createElement('tbody');
    
    // Flatten grouped data for the table
    groupedData.forEach(group => {
        group.contacts.forEach((contact, index) => {
            const row = document.createElement('tr');
            
            // Add company name only for the first row in each group
            if (index === 0) {
                headers.forEach((header, colIndex) => {
                    const td = document.createElement('td');
                    td.textContent = contact[header] || '';
                    row.appendChild(td);
                });
            } else {
                headers.forEach((header, colIndex) => {
                    const td = document.createElement('td');
                    // For column A (company), leave blank for non-first rows in a group
                    if (colIndex === 0) {
                        td.textContent = '';
                    } else {
                        td.textContent = contact[header] || '';
                    }
                    row.appendChild(td);
                });
            }
            
            // Add action button
            const actionCell = document.createElement('td');
            const viewButton = document.createElement('button');
            viewButton.className = 'btn tertiary';
            viewButton.innerHTML = '<i class="fas fa-eye"></i> View';
            viewButton.addEventListener('click', () => {
                state.currentContact = contact;
                populateContactDetails(contact);
                showSection(elements.contactInfoSection);
            });
            
            actionCell.appendChild(viewButton);
            row.appendChild(actionCell);
            
            tbody.appendChild(row);
        });
    });
    
    elements.contactsTable.appendChild(tbody);
}

// Event Handlers
function setupEventListeners() {
    // Load Sheets Button
    elements.loadSheets.addEventListener('click', async () => {
        const spreadsheetId = elements.spreadsheetId.value.trim();
        
        if (!spreadsheetId) {
            showToast('Please enter a Spreadsheet ID', 'warning');
            return;
        }
        
        try {
            state.spreadsheetId = spreadsheetId;
            const sheets = await api.fetchSheets(spreadsheetId);
            populateSheetSelect(sheets);
            elements.loadContacts.disabled = true;
            showToast('Sheets loaded successfully', 'success');
        } catch (error) {
            console.error('Error loading sheets:', error);
        }
    });
    
    // Sheet Select
    elements.sheetSelect.addEventListener('change', () => {
        const selectedSheet = elements.sheetSelect.value;
        state.selectedSheet = selectedSheet;
        elements.loadContacts.disabled = !selectedSheet;
    });
    
    // Load Contacts Button
    elements.loadContacts.addEventListener('click', async () => {
        try {
            const data = await api.fetchSheetData(state.spreadsheetId, state.selectedSheet);
            state.sheetData = data;
            showToast('Contacts loaded successfully', 'success');
        } catch (error) {
            console.error('Error loading contacts:', error);
        }
    });
    
    // Upload Template Button
    elements.uploadTemplate.addEventListener('click', async () => {
        const file = elements.templateFile.files[0];
        
        if (!file) {
            showToast('Please select a template file', 'warning');
            return;
        }
        
        if (!file.name.endsWith('.docx')) {
            showToast('Please select a Word document (.docx)', 'warning');
            return;
        }
        
        try {
            await api.uploadTemplate(file);
            const templates = await api.fetchTemplates();
            state.templates = templates;
            populateTemplatesList(templates);
            showToast('Template uploaded successfully', 'success');
            
            // Enable generate email button if a contact is selected
            if (state.currentContact) {
                elements.generateEmail.disabled = false;
            }
        } catch (error) {
            console.error('Error uploading template:', error);
        }
    });
    
    // Search by Line Button
    elements.searchByLine.addEventListener('click', async () => {
        const lineNumber = elements.lineNumber.value.trim();
        
        if (!lineNumber) {
            showToast('Please enter a line number', 'warning');
            return;
        }
        
        if (!state.spreadsheetId || !state.selectedSheet) {
            showToast('Please select a spreadsheet and sheet first', 'warning');
            return;
        }
        
        try {
            const contact = await api.fetchContactByLine(state.spreadsheetId, state.selectedSheet, lineNumber);
            state.currentContact = contact;
            populateContactDetails(contact);
            showSection(elements.contactInfoSection);
        } catch (error) {
            console.error('Error searching by line:', error);
            showContactNotFound();
        }
    });
    
    // Search by Name Button
    elements.searchByName.addEventListener('click', async () => {
        const name = elements.contactName.value.trim();
        
        if (!name) {
            showToast('Please enter a contact name', 'warning');
            return;
        }
        
        if (!state.spreadsheetId || !state.selectedSheet) {
            showToast('Please select a spreadsheet and sheet first', 'warning');
            return;
        }
        
        try {
            const contact = await api.fetchContactByName(state.spreadsheetId, state.selectedSheet, name);
            state.currentContact = contact;
            populateContactDetails(contact);
            showSection(elements.contactInfoSection);
        } catch (error) {
            console.error('Error searching by name:', error);
            showContactNotFound();
        }
    });
    
    // Generate Email Button
    elements.generateEmail.addEventListener('click', async () => {
        if (!state.currentContact) {
            showToast('Please select a contact first', 'warning');
            return;
        }
        
        if (!state.selectedTemplate) {
            showToast('Please select a template first', 'warning');
            return;
        }
        
        try {
            const emailData = await api.generateEmail(state.selectedTemplate, state.currentContact);
            
            // Populate email preview
            elements.emailTo.value = state.currentContact.Email || '';
            elements.emailSubject.value = emailData.subject || '';
            elements.emailContent.value = emailData.emailContent || '';
            
            state.emailContent = emailData.emailContent;
            
            // Show email preview section
            showSection(elements.emailPreviewSection);
        } catch (error) {
            console.error('Error generating email:', error);
        }
    });
    
    // View All Data Button
    elements.viewAllData.addEventListener('click', () => {
        if (!state.sheetData) {
            showToast('No data available. Please load contacts first.', 'warning');
            return;
        }
        
        populateDataTable(state.sheetData);
        showSection(elements.dataTableSection);
    });
    
    // Send Email Button
    elements.sendEmail.addEventListener('click', async () => {
        const to = elements.emailTo.value.trim();
        const subject = elements.emailSubject.value.trim();
        const content = elements.emailContent.value.trim();
        const from = elements.fromEmail.value.trim();
        
        if (!to) {
            showToast('Recipient email is required', 'warning');
            return;
        }
        
        if (!subject) {
            showToast('Subject is required', 'warning');
            return;
        }
        
        if (!content) {
            showToast('Email content is required', 'warning');
            return;
        }
        
        if (elements.emailMethod.value === 'direct') {
            try {
                await api.sendEmail(to, subject, content, from);
                showToast('Email sent successfully', 'success');
                showSection(elements.contactInfoSection);
            } catch (error) {
                console.error('Error sending email:', error);
            }
        } else {
            try {
                const mailtoUrl = await api.getOutlookUrl(to, subject, content);
                window.location.href = mailtoUrl;
            } catch (error) {
                console.error('Error opening Outlook:', error);
            }
        }
    });
    
    // Open in Outlook Button
    elements.openInOutlook.addEventListener('click', async () => {
        const to = elements.emailTo.value.trim();
        const subject = elements.emailSubject.value.trim();
        const body = elements.emailContent.value.trim();
        
        if (!to) {
            showToast('Recipient email is required', 'warning');
            return;
        }
        
        try {
            const mailtoUrl = await api.getOutlookUrl(to, subject, body);
            window.location.href = mailtoUrl;
        } catch (error) {
            console.error('Error opening Outlook:', error);
        }
    });
    
    // Back Buttons
    elements.backToContact.addEventListener('click', () => {
        showSection(elements.contactInfoSection);
    });
    
    elements.backFromTable.addEventListener('click', () => {
        showSection(elements.contactInfoSection);
    });
    
    // Load templates on page load
    window.addEventListener('load', async () => {
        try {
            const templates = await api.fetchTemplates();
            state.templates = templates;
            populateTemplatesList(templates);
        } catch (error) {
            console.error('Error loading templates:', error);
        }
    });
}

// Initialize the app
function init() {
    setupEventListeners();
    showSection(elements.contactInfoSection);
}

// Start the app
init();
