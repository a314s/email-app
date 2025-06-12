// DOM Elements
const elements = {
    // Excel File Configuration
    excelFile: document.getElementById('excelFile'),
    uploadExcel: document.getElementById('uploadExcel'),
    excelSelect: document.getElementById('excelSelect'),
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
    previewTo: document.getElementById('previewTo'),
    previewSubject: document.getElementById('previewSubject'),
    previewContent: document.getElementById('previewContent'),
    toggleRTL: document.getElementById('toggleRTL'),
    increaseFontSize: document.getElementById('increaseFontSize'),
    decreaseFontSize: document.getElementById('decreaseFontSize'),
    sendEmail: document.getElementById('sendEmail'),
    openInOutlook: document.getElementById('openInOutlook'),
    backToContact: document.getElementById('backToContact'),
    
    // Data Table
    dataTableSection: document.getElementById('dataTableSection'),
    contactsTable: document.getElementById('contactsTable'),
    backFromTable: document.getElementById('backFromTable'),
    
    // Data Entry Form
    dataEntrySection: document.getElementById('dataEntrySection'),
    dataEntryForm: document.getElementById('dataEntryForm'),
    dynamicFormFields: document.getElementById('dynamicFormFields'),
    submitForm: document.getElementById('submitForm'),
    cancelForm: document.getElementById('cancelForm'),
    
    // UI Elements
    toast: document.getElementById('toast'),
    loadingOverlay: document.getElementById('loadingOverlay'),
    
    // Calendar View
    calendarViewBtn: document.getElementById('calendarViewBtn'),
    companyCalendarBtn: document.getElementById('companyCalendarBtn'),
    calendarSection: document.getElementById('calendarSection'),
    backFromCalendar: document.getElementById('backFromCalendar'),
    prevMonth: document.getElementById('prevMonth'),
    nextMonth: document.getElementById('nextMonth'),
    todayBtn: document.getElementById('todayBtn'),
    calendarGrid: document.getElementById('calendarGrid'),
    followUpsList: document.getElementById('followUpsList'),
    currentMonthYear: document.getElementById('currentMonthYear'),
    companyFilter: document.getElementById('companyFilter'),
    settingsBtn: document.getElementById('settingsBtn')
};

// App State
const state = {
    excelFiles: [],
    selectedExcelFile: '',
    selectedSheet: '',
    sheetData: null,
    currentContact: null,
    templates: [],
    selectedTemplate: '',
    emailContent: '',
    contactNames: [], // Add this to store contact names for autocomplete
    currentYear: new Date().getFullYear(),
    currentMonth: new Date().getMonth() + 1,
    selectedCompany: '',
    modals: {
        emailDetails: null,
        followUpDetails: null,
        excelFollowUpDetails: null
    },
    currentFollowUp: null,
    settings: {
        emailSignature: localStorage.getItem('emailSignature') || '',
        defaultFromEmail: localStorage.getItem('defaultFromEmail') || '',
        defaultEmailMethod: localStorage.getItem('defaultEmailMethod') || 'direct'
    },
    excelColumnMapping: null
};

// API Functions
const api = {
    async uploadExcelFile(file) {
        showLoading(true);
        try {
            const formData = new FormData();
            formData.append('excelFile', file);
            
            const response = await fetch('/api/excel/upload', {
                method: 'POST',
                body: formData
            });
            
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to upload Excel file');
            }
            
            return data;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async fetchExcelFiles() {
        showLoading(true);
        try {
            const response = await fetch('/api/excel/files');
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to fetch Excel files');
            }
            
            return data.files;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async fetchSheets(filename) {
        showLoading(true);
        try {
            const response = await fetch(`/api/excel/${encodeURIComponent(filename)}/sheets`);
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
    
    async fetchSheetData(filename, sheetName) {
        showLoading(true);
        try {
            const response = await fetch(`/api/excel/${encodeURIComponent(filename)}/${encodeURIComponent(sheetName)}`);
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to fetch sheet data');
            }
            
            // Extract all contact names for autocomplete
            if (data.data && Array.isArray(data.data)) {
                const names = [];
                data.data.forEach(company => {
                    if (company.contacts && Array.isArray(company.contacts)) {
                        company.contacts.forEach(contact => {
                            // Add English name (usually in Name column)
                            if (contact['Name'] && !names.includes(contact['Name'])) {
                                names.push(contact['Name']);
                            }
                            
                            // Add Hebrew name (usually in column F)
                            // We'll check for Hebrew name in common column names
                            const hebrewName = contact['Hebrew Name'] || contact['שם'] || contact['F'] || contact['שם בעברית'];
                            if (hebrewName && !names.includes(hebrewName)) {
                                names.push(hebrewName);
                            }
                        });
                    }
                });
                state.contactNames = names;
            }
            
            return data;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async fetchContactByLine(filename, sheetName, lineNumber) {
        showLoading(true);
        try {
            const response = await fetch(`/api/contact/line/${encodeURIComponent(filename)}/${encodeURIComponent(sheetName)}/${lineNumber}`);
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
    
    async fetchContactByName(filename, sheetName, name) {
        showLoading(true);
        try {
            const response = await fetch(`/api/contact/name/${encodeURIComponent(filename)}/${encodeURIComponent(sheetName)}/${encodeURIComponent(name)}`);
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
    
    async improveEmailContent(content, contactData) {
        showLoading(true);
        try {
            console.log('Improving email content with contact data:', contactData);
            
            const response = await fetch('/api/improve-email', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ content, contactData })
            });
            
            const data = await response.json();
            console.log('API response:', data);
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to improve email content');
            }
            
            return data.improvedContent;
        } catch (error) {
            console.error('Error in improveEmailContent:', error);
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async sendEmail(to, subject, content, from, contactData = null) {
        showLoading(true);
        try {
            const requestBody = { 
                to, 
                subject, 
                content, 
                from 
            };
            
            // Add contact data if provided
            if (contactData) {
                requestBody.contactData = contactData;
            }
            
            // Add column mapping if available
            if (state.excelColumnMapping) {
                requestBody.columnMapping = state.excelColumnMapping;
            }
            
            const response = await fetch('/api/send-email', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(requestBody)
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
        // Format the body for mailto URL
        const formattedBody = body.replace(/\n/g, '%0D%0A');
        
        // Create the mailto URL
        const mailtoUrl = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(formattedBody)}`;
        
        return { success: true, mailtoUrl };
    },
    
    async updateExcelData(filename, sheetName, formData) {
        showLoading(true);
        try {
            const response = await fetch('/api/excel/update', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    filename,
                    sheetName,
                    formData
                })
            });
            
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to update Excel data');
            }
            
            return data;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async fetchCalendar(year, month) {
        showLoading(true);
        try {
            let url;
            if (state.selectedCompany) {
                url = `/api/calendar/${year}/${month}/${encodeURIComponent(state.selectedCompany)}`;
            } else {
                url = `/api/calendar/${year}/${month}`;
            }
            
            // Add column mapping if available
            if (state.excelColumnMapping) {
                const params = new URLSearchParams({
                    firstEmailColumn: state.excelColumnMapping.firstEmail,
                    secondEmailColumn: state.excelColumnMapping.secondEmail,
                    thirdEmailColumn: state.excelColumnMapping.thirdEmail
                });
                url += `?${params.toString()}`;
            }
            
            const response = await fetch(url);
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to fetch calendar data');
            }
            
            return data;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async completeFollowUp(followUpId, notes) {
        showLoading(true);
        try {
            const response = await fetch(`/api/follow-ups/${followUpId}/complete`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ notes })
            });
            
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to complete follow-up');
            }
            
            return data;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async fetchFollowUps() {
        showLoading(true);
        try {
            const response = await fetch('/api/follow-ups');
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to fetch follow-ups');
            }
            
            return data.followUps;
        } catch (error) {
            showToast(error.message, 'error');
            throw error;
        } finally {
            showLoading(false);
        }
    },
    
    async analyzeExcelColumns(filename, sheetName) {
        showLoading(true);
        try {
            const response = await fetch('/api/excel/analyze-columns', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    filename,
                    sheetName
                })
            });
            
            const data = await response.json();
            
            if (!data.success) {
                throw new Error(data.error || 'Failed to analyze Excel columns');
            }
            
            if (!data.analysis || !data.analysis.firstEmailColumn) {
                throw new Error('Invalid column analysis results');
            }
            
            return data.analysis;
        } catch (error) {
            console.error('Error analyzing Excel columns:', error);
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

function hideAllSections() {
    // Hide all sections
    elements.contactInfoSection.classList.add('hidden');
    elements.emailPreviewSection.classList.add('hidden');
    elements.dataTableSection.classList.add('hidden');
    elements.dataEntrySection.classList.add('hidden');
    elements.calendarSection.classList.add('hidden');
}

function showSection(section) {
    hideAllSections();
    
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
        
        // Add specific classes based on field type
        if (key.toLowerCase().includes('email')) {
            detailItem.classList.add('email-field');
        } else if (key.toLowerCase().includes('phone') || key.toLowerCase().includes('mobile')) {
            detailItem.classList.add('phone-field');
        } else if (key.toLowerCase().includes('company')) {
            detailItem.classList.add('company-field');
        }
        
        const label = document.createElement('div');
        label.className = 'label';
        label.textContent = key;
        
        const valueElement = document.createElement('div');
        valueElement.className = 'value';
        
        // Check if the value is empty or not
        if (value) {
            valueElement.classList.add('has-data');
            detailItem.classList.add('has-data');
            
            // Check if the value is an email
            if (key.toLowerCase().includes('email') && isValidEmail(value)) {
                const link = document.createElement('a');
                link.href = `mailto:${value}`;
                link.textContent = value;
                valueElement.appendChild(link);
            } 
            // Check if the value is a URL
            else if (isValidURL(value)) {
                const link = document.createElement('a');
                link.href = value.startsWith('http') ? value : `https://${value}`;
                link.textContent = value;
                link.target = '_blank';
                link.rel = 'noopener noreferrer';
                valueElement.appendChild(link);
            } 
            // Check if the value is a phone number
            else if (key.toLowerCase().includes('phone') || key.toLowerCase().includes('mobile') || key.toLowerCase().includes('tel')) {
                const link = document.createElement('a');
                link.href = `tel:${value.replace(/\D/g, '')}`;
                link.textContent = value;
                valueElement.appendChild(link);
            } 
            else {
                valueElement.textContent = value;
            }
        } else {
            valueElement.textContent = 'N/A';
            detailItem.classList.add('no-data');
        }
        
        detailItem.appendChild(label);
        detailItem.appendChild(valueElement);
        elements.contactDetailsGrid.appendChild(detailItem);
    }
    
    // Show contact details and hide "not found" message
    elements.contactDetails.classList.remove('hidden');
    elements.noContactFound.classList.add('hidden');
    
    // Enable generate email button if templates are available
    elements.generateEmail.disabled = state.templates.length === 0;
    
    // Add "Add New Data" button to contact actions
    const contactActions = document.querySelector('.contact-actions');
    if (contactActions) {
        // Check if the button already exists
        if (!document.getElementById('addNewData')) {
            const addNewDataBtn = document.createElement('button');
            addNewDataBtn.id = 'addNewData';
            addNewDataBtn.className = 'btn secondary';
            addNewDataBtn.innerHTML = '<i class="fas fa-plus"></i> Add New Data';
            addNewDataBtn.addEventListener('click', showDataEntryForm);
            contactActions.appendChild(addNewDataBtn);
        }
    }
}

function showContactNotFound() {
    elements.contactDetails.classList.add('hidden');
    elements.noContactFound.classList.remove('hidden');
}

function populateExcelSelect(files) {
    // Clear previous options
    elements.excelSelect.innerHTML = '';
    
    // Add default option
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = 'Select a file...';
    elements.excelSelect.appendChild(defaultOption);
    
    // Add options for each file
    files.forEach(file => {
        const option = document.createElement('option');
        option.value = file.filename;
        option.textContent = file.filename;
        elements.excelSelect.appendChild(option);
    });
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

// Helper function to check if a string is a valid email
function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

// Helper function to check if a string is a valid URL
function isValidURL(url) {
    // Simple check for common domain extensions or www
    return /\.(com|org|net|io|gov|edu|co|uk|de|fr|it|es|nl|ru|jp|cn|in|au|ca|br|mx)$/i.test(url) || 
           /^www\./i.test(url) || 
           /^http(s)?:\/\//i.test(url);
}

// Update email preview as user types
function updateEmailPreview() {
    elements.previewTo.textContent = elements.emailTo.value;
    elements.previewSubject.textContent = elements.emailSubject.value;
    
    // Convert plain text to HTML with line breaks
    const content = elements.emailContent.value
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/\n/g, '<br>');
    
    elements.previewContent.innerHTML = content;
}

// Toggle RTL/LTR direction
function toggleDirection() {
    const currentDir = elements.emailContent.getAttribute('dir');
    const newDir = currentDir === 'ltr' ? 'rtl' : 'ltr';
    
    // Update textarea direction
    elements.emailContent.setAttribute('dir', newDir);
    
    // Update preview direction
    elements.previewContent.setAttribute('dir', newDir);
    
    // If switching to RTL and we have a contact with a Hebrew name, use it in the greeting
    if (newDir === 'rtl' && state.currentContact) {
        const hebrewName = state.currentContact['Hebrew Name'] || 
                          state.currentContact['שם'] || 
                          state.currentContact['F'] || 
                          state.currentContact['שם בעברית'];
        
        if (hebrewName) {
            // Replace the English name with Hebrew name in the email content
            const englishName = state.currentContact['Name'] || '';
            if (englishName && elements.emailContent.value.includes(englishName)) {
                elements.emailContent.value = elements.emailContent.value.replace(
                    new RegExp(englishName, 'g'), 
                    hebrewName
                );
                // Update the preview
                updateEmailPreview();
            }
        }
    }
    
    showToast(`Text direction changed to ${newDir.toUpperCase()}`, 'info');
}

// Change font size
function changeFontSize(increase = true) {
    const textarea = elements.emailContent;
    const preview = elements.previewContent;
    
    // Get current font size
    const currentSize = parseInt(window.getComputedStyle(textarea).fontSize);
    
    // Calculate new size (min: 10px, max: 24px)
    let newSize;
    if (increase) {
        newSize = Math.min(currentSize + 1, 24);
    } else {
        newSize = Math.max(currentSize - 1, 10);
    }
    
    // Apply new size
    textarea.style.fontSize = `${newSize}px`;
    preview.style.fontSize = `${newSize}px`;
}

// Setup autocomplete for contact name search
function setupContactNameAutocomplete() {
    const input = elements.contactName;
    let currentFocus;
    
    // Execute when someone writes in the text field
    input.addEventListener("input", function(e) {
        let a, b, i, val = this.value;
        // Close any already open lists
        closeAllLists();
        if (!val) { return false; }
        currentFocus = -1;
        
        // Create a DIV element that will contain the autocomplete items
        a = document.createElement("DIV");
        a.setAttribute("id", this.id + "autocomplete-list");
        a.setAttribute("class", "autocomplete-items");
        this.parentNode.appendChild(a);
        
        // For each item in the array
        for (i = 0; i < state.contactNames.length; i++) {
            // Check if the item starts with the same letters as the text field value
            if (state.contactNames[i].substr(0, val.length).toUpperCase() == val.toUpperCase()) {
                // Create a DIV element for each matching element
                b = document.createElement("DIV");
                // Make the matching letters bold
                b.innerHTML = "<strong>" + state.contactNames[i].substr(0, val.length) + "</strong>";
                b.innerHTML += state.contactNames[i].substr(val.length);
                // Insert a input field that will hold the current array item's value
                b.innerHTML += "<input type='hidden' value='" + state.contactNames[i] + "'>";
                
                // Execute a function when someone clicks on the item value
                b.addEventListener("click", function(e) {
                    // Insert the value for the autocomplete text field
                    input.value = this.getElementsByTagName("input")[0].value;
                    // Close the list of autocompleted values
                    closeAllLists();
                });
                a.appendChild(b);
            }
        }
    });
    
    // Execute a function when pressing a key on the keyboard
    input.addEventListener("keydown", function(e) {
        let x = document.getElementById(this.id + "autocomplete-list");
        if (x) x = x.getElementsByTagName("div");
        if (e.keyCode == 40) { // DOWN key
            currentFocus++;
            addActive(x);
        } else if (e.keyCode == 38) { // UP key
            currentFocus--;
            addActive(x);
        } else if (e.keyCode == 13) { // ENTER key
            e.preventDefault();
            if (currentFocus > -1) {
                if (x) x[currentFocus].click();
            }
        }
    });
    
    function addActive(x) {
        if (!x) return false;
        removeActive(x);
        if (currentFocus >= x.length) currentFocus = 0;
        if (currentFocus < 0) currentFocus = (x.length - 1);
        x[currentFocus].classList.add("autocomplete-active");
    }
    
    function removeActive(x) {
        for (let i = 0; i < x.length; i++) {
            x[i].classList.remove("autocomplete-active");
        }
    }
    
    function closeAllLists(elmnt) {
        const x = document.getElementsByClassName("autocomplete-items");
        for (let i = 0; i < x.length; i++) {
            if (elmnt != x[i] && elmnt != input) {
                x[i].parentNode.removeChild(x[i]);
            }
        }
    }
    
    // Close all lists when clicking elsewhere
    document.addEventListener("click", function (e) {
        closeAllLists(e.target);
    });
}

// Populate the dynamic form fields based on Excel headers
function generateFormFields(headers) {
    elements.dynamicFormFields.innerHTML = '';
    
    headers.forEach(header => {
        if (header) { // Skip empty headers
            const formGroup = document.createElement('div');
            formGroup.className = 'form-group';
            
            const label = document.createElement('label');
            label.setAttribute('for', `form-${header}`);
            label.textContent = header;
            
            const input = document.createElement('input');
            input.type = 'text';
            input.id = `form-${header}`;
            input.name = header;
            input.placeholder = `Enter ${header}`;
            
            formGroup.appendChild(label);
            formGroup.appendChild(input);
            elements.dynamicFormFields.appendChild(formGroup);
        }
    });
}

// Show the data entry form
function showDataEntryForm() {
    if (!state.selectedExcelFile || !state.selectedSheet) {
        showToast('Please select an Excel file and sheet first', 'warning');
        return;
    }
    
    if (!state.sheetData || !state.sheetData.headers) {
        showToast('Please load contacts first to get sheet structure', 'warning');
        return;
    }
    
    // Generate form fields based on headers
    generateFormFields(state.sheetData.headers);
    
    // Show the form section
    showSection(elements.dataEntrySection);
}

// Handle form submission
function handleFormSubmit(event) {
    event.preventDefault();
    
    if (!state.selectedExcelFile || !state.selectedSheet) {
        showToast('Please select an Excel file and sheet first', 'warning');
        return;
    }
    
    // Collect form data
    const formData = {};
    const formElements = elements.dataEntryForm.elements;
    
    for (let i = 0; i < formElements.length; i++) {
        const element = formElements[i];
        if (element.name && element.name !== '') {
            formData[element.name] = element.value;
        }
    }
    
    // Submit the form data
    api.updateExcelData(state.selectedExcelFile, state.selectedSheet, formData)
        .then(result => {
            showToast('Data added successfully', 'success');
            
            // Reset form
            elements.dataEntryForm.reset();
            
            // Refresh the data
            return api.fetchSheetData(state.selectedExcelFile, state.selectedSheet);
        })
        .then(data => {
            state.sheetData = data;
            showToast('Data refreshed successfully', 'success');
            
            // Go back to contact info section
            showSection(elements.contactInfoSection);
        })
        .catch(error => {
            console.error('Error submitting form:', error);
        });
}

// Event Handlers
document.addEventListener('DOMContentLoaded', function() {
    // Upload Excel Button
    elements.uploadExcel.addEventListener('click', async () => {
        const file = elements.excelFile.files[0];
        
        if (!file) {
            showToast('Please select an Excel file', 'warning');
            return;
        }
        
        if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
            showToast('Please select an Excel file (.xlsx or .xls)', 'warning');
            return;
        }
        
        try {
            const result = await api.uploadExcelFile(file);
            showToast('Excel file uploaded successfully', 'success');
            
            // Update the Excel files list
            const files = await api.fetchExcelFiles();
            state.excelFiles = files;
            populateExcelSelect(files);
            
            // Set the selected file
            state.selectedExcelFile = result.filename;
            elements.excelSelect.value = result.filename;
            
            // Load sheets for the uploaded file
            const sheets = result.sheets;
            populateSheetSelect(sheets);
            
            // Reset sheet selection
            state.selectedSheet = '';
            elements.loadContacts.disabled = true;
        } catch (error) {
            console.error('Error uploading Excel file:', error);
        }
    });
    
    // Excel Select
    elements.excelSelect.addEventListener('change', async () => {
        const selectedFile = elements.excelSelect.value;
        
        if (!selectedFile) {
            elements.sheetSelect.disabled = true;
            elements.sheetSelect.innerHTML = '<option value="">Select a sheet...</option>';
            elements.loadContacts.disabled = true;
            return;
        }
        
        state.selectedExcelFile = selectedFile;
        
        try {
            const sheets = await api.fetchSheets(selectedFile);
            populateSheetSelect(sheets);
            elements.loadContacts.disabled = true;
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
        if (!state.selectedExcelFile || !state.selectedSheet) {
            showToast('Please select an Excel file and sheet', 'warning');
            return;
        }
        
        try {
            // First, analyze the columns
            const columnAnalysis = await api.analyzeExcelColumns(state.selectedExcelFile, state.selectedSheet);
            
            // Store the column mapping in state
            state.excelColumnMapping = {
                firstEmail: columnAnalysis.firstEmailColumn,
                secondEmail: columnAnalysis.secondEmailColumn,
                thirdEmail: columnAnalysis.thirdEmailColumn
            };
            
            // Show the analysis results to the user
            showToast(`Identified email columns with ${columnAnalysis.confidence} confidence: ${columnAnalysis.explanation}`, 'info');
            
            // Load the sheet data
            const data = await api.fetchSheetData(state.selectedExcelFile, state.selectedSheet);
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
        
        if (!state.selectedExcelFile || !state.selectedSheet) {
            showToast('Please select an Excel file and sheet first', 'warning');
            return;
        }
        
        try {
            const contact = await api.fetchContactByLine(state.selectedExcelFile, state.selectedSheet, lineNumber);
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
        
        if (!state.selectedExcelFile || !state.selectedSheet) {
            showToast('Please select an Excel file and sheet first', 'warning');
            return;
        }
        
        try {
            const contact = await api.fetchContactByName(state.selectedExcelFile, state.selectedSheet, name);
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
            // Check for email in different possible field names
            const emailAddress = state.currentContact.Email || 
                                state.currentContact.email || 
                                state.currentContact['E-mail'] || 
                                state.currentContact['Email Address'] || 
                                '';
            
            elements.emailTo.value = emailAddress;
            elements.emailSubject.value = emailData.subject || "What can navitech Aid do for you";
            
            // Add signature to email content if it exists
            let content = emailData.emailContent || '';
            if (state.settings.emailSignature) {
                content += '\n\n' + state.settings.emailSignature;
            }
            elements.emailContent.value = content;
            
            state.emailContent = emailData.emailContent;
            
            // Update the visual preview
            updateEmailPreview();
            
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
                // Prepare contact data with Excel file information
                let contactData = null;
                if (state.currentContact) {
                    contactData = {
                        ...state.currentContact,
                        excelFile: state.selectedExcelFile,
                        sheetName: state.selectedSheet
                    };
                }
                
                await api.sendEmail(to, subject, content, from, contactData);
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
            const { mailtoUrl } = await api.getOutlookUrl(to, subject, body);
            window.open(mailtoUrl, '_blank');
        } catch (error) {
            console.error('Error opening Outlook:', error);
            showToast('Failed to open email client', 'error');
        }
    });
    
    // Toggle RTL/LTR button
    elements.toggleRTL.addEventListener('click', toggleDirection);
    
    // Font size buttons
    elements.increaseFontSize.addEventListener('click', () => changeFontSize(true));
    elements.decreaseFontSize.addEventListener('click', () => changeFontSize(false));
    
    // Back Buttons
    elements.backToContact.addEventListener('click', () => {
        showSection(elements.contactInfoSection);
    });
    
    elements.backFromTable.addEventListener('click', () => {
        showSection(elements.contactInfoSection);
    });
    
    // Update preview as user types in email fields
    elements.emailTo.addEventListener('input', updateEmailPreview);
    elements.emailSubject.addEventListener('input', updateEmailPreview);
    elements.emailContent.addEventListener('input', updateEmailPreview);
    
    // Add a button to improve email content using Anthropic API
    const improveEmailBtn = document.createElement('button');
    improveEmailBtn.id = 'improveEmail';
    improveEmailBtn.className = 'btn secondary';
    improveEmailBtn.innerHTML = '<i class="fas fa-magic"></i> Improve Email';
    improveEmailBtn.addEventListener('click', async () => {
        const content = elements.emailContent.value.trim();
        
        if (!content) {
            showToast('Email content is required', 'warning');
            return;
        }
        
        try {
            const improvedContent = await api.improveEmailContent(content, state.currentContact);
            elements.emailContent.value = improvedContent;
            updateEmailPreview();
            showToast('Email content improved', 'success');
        } catch (error) {
            console.error('Error improving email content:', error);
            showToast('Failed to improve email content', 'error');
        }
    });
    
    // Insert the improve email button after the textarea controls
    const textareaControls = document.querySelector('.textarea-controls');
    textareaControls.appendChild(improveEmailBtn);
    
    // Cancel Form Button
    elements.cancelForm.addEventListener('click', () => {
        elements.dataEntryForm.reset();
        showSection(elements.contactInfoSection);
    });
    
    // Form Submit
    elements.dataEntryForm.addEventListener('submit', handleFormSubmit);
    
    // Calendar View Button
    elements.calendarViewBtn.addEventListener('click', () => {
        showSection(elements.calendarSection);
        renderCalendar(new Date().getFullYear(), new Date().getMonth() + 1);
        updateFollowUpsList();
    });
    
    // Company Calendar Button
    elements.companyCalendarBtn.addEventListener('click', showCompanyCalendarView);

    // Settings Button
    elements.settingsBtn.addEventListener('click', () => {
        const modal = document.getElementById('settingsModal');
        modal.style.display = 'block';

        // Load current settings
        document.getElementById('defaultFromEmail').value = state.settings.defaultFromEmail;
        document.getElementById('defaultEmailMethod').value = state.settings.defaultEmailMethod;
        document.getElementById('emailSignature').value = state.settings.emailSignature;

        // Close modal when clicking the X
        modal.querySelector('.close-modal').addEventListener('click', () => {
            modal.style.display = 'none';
        });

        // Close modal when clicking outside
        window.addEventListener('click', (e) => {
            if (e.target === modal) {
                modal.style.display = 'none';
            }
        });
    });

    // Save Settings Button
    document.getElementById('saveSettings').addEventListener('click', () => {
        // Update settings in state and localStorage
        state.settings.defaultFromEmail = document.getElementById('defaultFromEmail').value;
        state.settings.defaultEmailMethod = document.getElementById('defaultEmailMethod').value;
        state.settings.emailSignature = document.getElementById('emailSignature').value;

        localStorage.setItem('defaultFromEmail', state.settings.defaultFromEmail);
        localStorage.setItem('defaultEmailMethod', state.settings.defaultEmailMethod);
        localStorage.setItem('emailSignature', state.settings.emailSignature);

        // Update UI with new settings
        elements.fromEmail.value = state.settings.defaultFromEmail;
        elements.emailMethod.value = state.settings.defaultEmailMethod;

        // Close the modal
        document.getElementById('settingsModal').style.display = 'none';
        showToast('Settings saved successfully', 'success');
    });
    
    // Back from Calendar Button
    elements.backFromCalendar.addEventListener('click', () => {
        showSection(elements.contactInfoSection);
    });
    
    // Previous Month Button
    elements.prevMonth.addEventListener('click', () => {
        const currentDate = new Date(elements.currentMonthYear.textContent);
        currentDate.setMonth(currentDate.getMonth() - 1);
        renderCalendar(currentDate.getFullYear(), currentDate.getMonth() + 1);
    });
    
    // Next Month Button
    elements.nextMonth.addEventListener('click', () => {
        const currentDate = new Date(elements.currentMonthYear.textContent);
        currentDate.setMonth(currentDate.getMonth() + 1);
        renderCalendar(currentDate.getFullYear(), currentDate.getMonth() + 1);
    });
    
    // Today Button
    elements.todayBtn.addEventListener('click', () => {
        renderCalendar(new Date().getFullYear(), new Date().getMonth() + 1);
    });
    
    // Company Filter
    elements.companyFilter.addEventListener('change', function() {
        state.selectedCompany = this.value;
        renderCalendar(state.currentYear, state.currentMonth);
    });
});

// Initialize the app
function init() {
    setupContactNameAutocomplete();
    
    // Apply saved settings
    if (state.settings.defaultFromEmail) {
        elements.fromEmail.value = state.settings.defaultFromEmail;
    }
    if (state.settings.defaultEmailMethod) {
        elements.emailMethod.value = state.settings.defaultEmailMethod;
    }
    
    api.fetchExcelFiles()
        .then(files => {
            state.excelFiles = files;
            populateExcelSelect(files);
        })
        .catch(error => console.error('Error fetching Excel files:', error));
    
    api.fetchTemplates()
        .then(templates => {
            state.templates = templates;
            populateTemplatesList(templates);
        })
        .catch(error => console.error('Error fetching templates:', error));
    
    // Load companies for the filter
    loadCompanies();
}

// Start the app
init();

// Calendar View Functionality
let currentDate = new Date();
let selectedDate = null;

// Render the calendar for a specific month
function renderCalendar(year, month) {
    // Update the month/year display
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 
                        'July', 'August', 'September', 'October', 'November', 'December'];
    elements.currentMonthYear.textContent = `${monthNames[month-1]} ${year}`;
    
    // Get calendar data from the server
    api.fetchCalendar(year, month)
        .then(data => {
            if (data.success) {
                // Generate calendar grid
                generateCalendarGrid(year, month, data.calendarData);
            } else {
                console.error('Error fetching calendar data:', data.error);
            }
        })
        .catch(error => {
            console.error('Error fetching calendar data:', error);
        });
}

// Generate the calendar grid
function generateCalendarGrid(year, month, calendarData) {
    elements.calendarGrid.innerHTML = '';
    
    // Get the first day of the month (0 = Sunday, 1 = Monday, etc.)
    const firstDay = new Date(year, month - 1, 1).getDay();
    
    // Get the number of days in the month
    const daysInMonth = new Date(year, month, 0).getDate();
    
    // Get the number of days in the previous month
    const daysInPrevMonth = new Date(year, month - 1, 0).getDate();
    
    // Get today's date for highlighting
    const today = new Date();
    const isCurrentMonth = today.getFullYear() === year && today.getMonth() === month - 1;
    
    // Calculate the number of rows needed (maximum 6)
    const rows = Math.ceil((firstDay + daysInMonth) / 7);
    
    // Generate days from previous month
    for (let i = 0; i < firstDay; i++) {
        const day = daysInPrevMonth - firstDay + i + 1;
        const dayElement = createDayElement(day, 'other-month', {});
        elements.calendarGrid.appendChild(dayElement);
    }
    
    // Generate days of the current month
    for (let day = 1; day <= daysInMonth; day++) {
        const isToday = isCurrentMonth && today.getDate() === day;
        const dayData = calendarData[day] || { emails: [], followUps: [] };
        
        const dayElement = createDayElement(day, isToday ? 'today' : '', dayData);
        elements.calendarGrid.appendChild(dayElement);
    }
    
    // Calculate the remaining cells to fill the grid
    const remainingCells = rows * 7 - (firstDay + daysInMonth);
    
    // Generate days from next month
    for (let day = 1; day <= remainingCells; day++) {
        const dayElement = createDayElement(day, 'other-month', {});
        elements.calendarGrid.appendChild(dayElement);
    }
}

// Create a day element for the calendar
function createDayElement(day, className, dayData) {
    const dayElement = document.createElement('div');
    dayElement.className = `calendar-day ${className}`;
    
    // Add day number
    const dayNumber = document.createElement('div');
    dayNumber.className = 'day-number';
    dayNumber.textContent = day;
    dayElement.appendChild(dayNumber);
    
    // Add events container
    const eventsContainer = document.createElement('div');
    eventsContainer.className = 'day-events';
    
    // Add emails
    if (dayData.emails && dayData.emails.length > 0) {
        dayData.emails.forEach(email => {
            const emailElement = document.createElement('div');
            emailElement.classList.add('day-email');
            emailElement.textContent = `Email: ${email.recipient.split('@')[0]}`;
            emailElement.dataset.id = email.id;
            emailElement.dataset.type = 'email';
            emailElement.addEventListener('click', () => showEmailDetails(email));
            eventsContainer.appendChild(emailElement);
        });
    }
    
    // Add Excel email dates
    if (dayData.excelFollowUps && dayData.excelFollowUps.length > 0) {
        dayData.excelFollowUps.forEach(excelData => {
            const excelElement = document.createElement('div');
            excelElement.classList.add('day-excel-follow-up');
            excelElement.textContent = `Excel Email #${excelData.followUpNumber}: ${excelData.contact_name}`;
            excelElement.dataset.id = excelData.id;
            excelElement.dataset.type = 'excelFollowUp';
            excelElement.addEventListener('click', () => showExcelFollowUpDetails(excelData));
            eventsContainer.appendChild(excelElement);
        });
    }
    
    // Add follow-ups
    if (dayData.followUps && dayData.followUps.length > 0) {
        dayData.followUps.forEach(followUp => {
            const followUpElement = document.createElement('div');
            followUpElement.classList.add('day-follow-up');
            followUpElement.textContent = `Follow-up #${followUp.follow_up_number}: ${followUp.recipient.split('@')[0]}`;
            followUpElement.dataset.id = followUp.id;
            followUpElement.dataset.type = 'followUp';
            followUpElement.addEventListener('click', () => showFollowUpDetails(followUp));
            eventsContainer.appendChild(followUpElement);
            
            // Also add to the follow-ups list if it's pending
            if (followUp.status === 'pending') {
                addFollowUpToList(followUp);
            }
        });
    }
    
    dayElement.appendChild(eventsContainer);
    return dayElement;
}

// Load companies for the filter
function loadCompanies() {
    fetch('/api/companies')
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                const companyFilter = elements.companyFilter;
                // Clear existing options except the first one (All Companies)
                while (companyFilter.options.length > 1) {
                    companyFilter.remove(1);
                }
                
                // Add companies to the filter
                data.companies.forEach(company => {
                    const option = document.createElement('option');
                    option.value = company;
                    option.textContent = company;
                    companyFilter.appendChild(option);
                });
            }
        })
        .catch(error => {
            console.error('Error loading companies:', error);
            showToast('Failed to load companies', 'error');
        });
}

// Update the follow-ups list
function updateFollowUpsList() {
    // Fetch pending follow-ups
    fetch('/api/follow-ups')
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                elements.followUpsList.innerHTML = '';
                
                if (data.followUps.length === 0) {
                    elements.followUpsList.innerHTML = '<div class="no-follow-ups">No pending follow-ups</div>';
                    return;
                }
                
                // Filter by company if selected
                let followUps = data.followUps;
                if (state.selectedCompany) {
                    followUps = followUps.filter(followUp => followUp.company === state.selectedCompany);
                }
                
                // Sort by date
                followUps.sort((a, b) => new Date(a.follow_up_date) - new Date(b.follow_up_date));
                
                // Add to the list
                followUps.forEach(followUp => {
                    addFollowUpToList(followUp);
                });
            }
        })
        .catch(error => {
            console.error('Error fetching follow-ups:', error);
            showToast('Failed to load follow-ups', 'error');
        });
}

// Add a follow-up to the list
function addFollowUpToList(followUp) {
    const followUpItem = document.createElement('div');
    followUpItem.classList.add('follow-up-item');
    
    // Format the date
    const followUpDate = new Date(followUp.follow_up_date);
    const formattedDate = followUpDate.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
    });
    
    followUpItem.innerHTML = `
        <div class="follow-up-info">
            <div><strong>${followUp.recipient}</strong></div>
            <div>${followUp.subject}</div>
            <div>Follow-up #${followUp.follow_up_number} - ${formattedDate}</div>
        </div>
        <div class="follow-up-actions">
            <button class="btn btn-sm btn-success complete-btn">Complete</button>
        </div>
    `;
    
    // Add click event to show details
    followUpItem.addEventListener('click', (e) => {
        if (!e.target.classList.contains('complete-btn')) {
            showFollowUpDetails(followUp);
        }
    });
    
    // Add click event to complete button
    const completeBtn = followUpItem.querySelector('.complete-btn');
    completeBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        completeFollowUp(followUp.id);
    });
    
    elements.followUpsList.appendChild(followUpItem);
}

// Complete a follow-up
function completeFollowUp(followUpId) {
    fetch(`/api/follow-ups/${followUpId}/complete`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ notes: 'Completed from calendar view' })
    })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                showToast('Follow-up marked as completed', 'success');
                // Refresh the calendar
                renderCalendar(state.currentYear, state.currentMonth);
            } else {
                showToast(`Failed to complete follow-up: ${data.error}`, 'error');
            }
        })
        .catch(error => {
            console.error('Error completing follow-up:', error);
            showToast('Failed to complete follow-up', 'error');
        });
}

// Show email details in a modal
function showEmailDetails(email) {
    const modal = document.getElementById('emailDetailsModal');
    const content = document.getElementById('emailDetailsContent');
    
    // Format the date
    const sentDate = new Date(email.sent_date);
    const formattedDate = sentDate.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    });
    
    content.innerHTML = `
        <div class="detail-row">
            <div class="detail-label">Recipient:</div>
            <div class="detail-value">${email.recipient}</div>
        </div>
        <div class="detail-row">
            <div class="detail-label">Subject:</div>
            <div class="detail-value">${email.subject}</div>
        </div>
        <div class="detail-row">
            <div class="detail-label">Sent Date:</div>
            <div class="detail-value">${formattedDate}</div>
        </div>
        <div class="detail-row">
            <div class="detail-label">Company:</div>
            <div class="detail-value">${email.company || 'N/A'}</div>
        </div>
        <div class="detail-row">
            <div class="detail-label">Contact:</div>
            <div class="detail-value">${email.contact_name || 'N/A'}</div>
        </div>
        <hr>
        <div class="detail-row">
            <div class="detail-value">${email.content}</div>
        </div>
    `;
    
    // Show the modal
    modal.style.display = 'block';
    
    // Close modal when clicking the X
    modal.querySelector('.close-modal').addEventListener('click', () => {
        modal.style.display = 'none';
    });
    
    // Close modal when clicking outside
    window.addEventListener('click', (e) => {
        if (e.target === modal) {
            modal.style.display = 'none';
        }
    });
}

// Show follow-up details in a modal
function showFollowUpDetails(followUp) {
    const modal = document.getElementById('followUpDetailsModal');
    const content = document.getElementById('followUpDetailsContent');
    
    // Store the current follow-up for the complete button
    state.currentFollowUp = followUp;
    
    // Format the date
    const followUpDate = new Date(followUp.follow_up_date);
    const formattedDate = followUpDate.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });
    
    content.innerHTML = `
        <div class="detail-row">
            <div class="detail-label">Recipient:</div>
            <div class="detail-value">${followUp.recipient}</div>
        </div>
        <div class="detail-row">
            <div class="detail-label">Subject:</div>
            <div class="detail-value">${followUp.subject}</div>
        </div>
        <div class="detail-row">
            <div class="detail-label">Follow-up Date:</div>
            <div class="detail-value">${formattedDate}</div>
        </div>
        <div class="detail-row">
            <div class="detail-label">Follow-up #:</div>
            <div class="detail-value">${followUp.follow_up_number}</div>
        </div>
        <div class="detail-row">
            <div class="detail-label">Status:</div>
            <div class="detail-value">${followUp.status}</div>
        </div>
        <div class="detail-row">
            <div class="detail-label">Company:</div>
            <div class="detail-value">${followUp.company || 'N/A'}</div>
        </div>
        <div class="detail-row">
            <div class="detail-label">Contact:</div>
            <div class="detail-value">${followUp.contact_name || 'N/A'}</div>
        </div>
    `;
    
    // Show the modal
    modal.style.display = 'block';
    
    // Close modal when clicking the X
    modal.querySelector('.close-modal').addEventListener('click', () => {
        modal.style.display = 'none';
    });
    
    // Close modal when clicking outside
    window.addEventListener('click', (e) => {
        if (e.target === modal) {
            modal.style.display = 'none';
        }
    });
    
    // Complete button event
    document.getElementById('completeFollowUpBtn').addEventListener('click', () => {
        completeFollowUp(followUp.id);
        modal.style.display = 'none';
    });
}

// Show Excel follow-up details in a modal
function showExcelFollowUpDetails(excelData) {
    const modal = document.getElementById('emailDetailsModal');
    const content = document.getElementById('emailDetailsContent');
    
    content.innerHTML = `
        <div class="detail-row">
            <div class="detail-label">Contact:</div>
            <div class="detail-value">${excelData.contact_name}</div>
        </div>
        <div class="detail-row">
            <div class="detail-label">Company:</div>
            <div class="detail-value">${excelData.company}</div>
        </div>
        <div class="detail-row">
            <div class="detail-label">Email:</div>
            <div class="detail-value">${excelData.recipient}</div>
        </div>
        <div class="detail-row">
            <div class="detail-label">Follow-up Number:</div>
            <div class="detail-value">${excelData.followUpNumber}</div>
        </div>
        <div class="detail-row">
            <div class="detail-label">Date:</div>
            <div class="detail-value">${new Date(excelData.followUpDate).toLocaleDateString()}</div>
        </div>
        <div class="detail-row">
            <div class="detail-label">Status:</div>
            <div class="detail-value">${excelData.status || 'Not set'}</div>
        </div>
    `;
    
    modal.style.display = 'block';
}

// Show company calendar view
function showCompanyCalendarView() {
    hideAllSections();
    elements.calendarSection.classList.remove('hidden');
    
    // Load companies for the filter
    loadCompanies();
    
    // Show company filter dropdown and prompt user to select
    elements.companyFilter.value = '';
    const companyFilterContainer = document.querySelector('.company-filter');
    companyFilterContainer.style.display = 'flex';
    
    // Render the current month
    renderCalendar(state.currentYear, state.currentMonth);
    
    // Show a toast message to guide the user
    showToast('Please select a company from the dropdown to view its calendar', 'info');
}

// Update the importFollowUpDataFromExcel function to use the identified columns
function importFollowUpDataFromExcel(excelFile, sheetName, contactData) {
    return new Promise((resolve, reject) => {
        try {
            // Check if we have follow-up data in the Excel
            if (!contactData) {
                resolve(null);
                return;
            }
            
            // Use the identified columns from the analysis
            const firstEmailDate = contactData[state.excelColumnMapping.firstEmail] || null;
            const firstEmailStatus = contactData[`${state.excelColumnMapping.firstEmail} Status`] || null;
            const secondEmailDate = contactData[state.excelColumnMapping.secondEmail] || null;
            const secondEmailStatus = contactData[`${state.excelColumnMapping.secondEmail} Status`] || null;
            const thirdEmailDate = contactData[state.excelColumnMapping.thirdEmail] || null;
            const thirdEmailStatus = contactData[`${state.excelColumnMapping.thirdEmail} Status`] || null;
            
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
