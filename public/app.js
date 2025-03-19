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
    
    // UI Elements
    toast: document.getElementById('toast'),
    loadingOverlay: document.getElementById('loadingOverlay')
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
    contactNames: [] // Add this to store contact names for autocomplete
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
                            // Assuming the name is in column E (index 4)
                            if (contact['Name'] && !names.includes(contact['Name'])) {
                                names.push(contact['Name']);
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
        // Format the body for mailto URL
        const formattedBody = body.replace(/\n/g, '%0D%0A');
        
        // Create the mailto URL
        const mailtoUrl = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(formattedBody)}`;
        
        return { success: true, mailtoUrl };
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
        
        // Check if the value is empty or not
        if (value) {
            valueElement.classList.add('has-data');
            
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

// Event Handlers
function setupEventListeners() {
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
            elements.emailContent.value = emailData.emailContent || '';
            
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
}

// Initialize the app
function init() {
    setupEventListeners();
    setupContactNameAutocomplete();
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
}

// Start the app
init();
