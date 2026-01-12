// Global variables
let msalInstance;
let currentUser = null;
let isAdmin = false;
let accessToken = null;

// Initialize MSAL
function initializeMsal() {
    const msalConfig = {
        auth: CONFIG.auth,
        cache: CONFIG.cache
    };
    
    msalInstance = new msal.PublicClientApplication(msalConfig);
    
    // Handle redirect response
    msalInstance.handleRedirectPromise()
        .then(response => {
            if (response) {
                currentUser = response.account;
                handleAuthentication();
            } else {
                checkExistingAuth();
            }
        })
        .catch(error => {
            console.error('Authentication error:', error);
            showAlert('submitAlert', 'Authentication failed. Please try again.', 'error');
        });
}

// Check if user is already authenticated
function checkExistingAuth() {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        currentUser = accounts[0];
        handleAuthentication();
    }
}

// Sign in
async function signIn() {
    try {
        await msalInstance.loginRedirect(loginRequest);
    } catch (error) {
        console.error('Sign in error:', error);
        showAlert('submitAlert', 'Sign in failed. Please try again.', 'error');
    }
}

// Sign out
async function signOut() {
    try {
        await msalInstance.logoutRedirect();
    } catch (error) {
        console.error('Sign out error:', error);
    }
}

// Handle authentication
async function handleAuthentication() {
    if (currentUser) {
        // Show app content
        document.getElementById('loginSection').style.display = 'none';
        document.getElementById('appContent').style.display = 'block';
        
        // Display user info
        document.getElementById('userInfo').innerHTML = `
            👤 ${currentUser.name} (${currentUser.username})
            <button class="btn btn-secondary" style="margin-left: 15px; padding: 5px 15px; font-size: 12px;" onclick="signOut()">
                Sign Out
            </button>
        `;
        
        // Check if user is admin
        isAdmin = CONFIG.admins.includes(currentUser.username.toLowerCase());
        
        if (isAdmin) {
            // Show admin tabs
            document.querySelectorAll('.admin-only').forEach(el => {
                el.style.display = 'block';
            });
        }
        
        // Get access token
        await getAccessToken();
        
        // Pre-fill form with user details
        document.getElementById('requestedBy').value = currentUser.name;
    }
}

// Get access token
async function getAccessToken() {
    const account = msalInstance.getAllAccounts()[0];
    
    if (!account) {
        console.error('No account found');
        return null;
    }
    
    const silentRequest = {
        scopes: CONFIG.scopes.graph,
        account: account
    };
    
    try {
        const response = await msalInstance.acquireTokenSilent(silentRequest);
        accessToken = response.accessToken;
        return accessToken;
    } catch (error) {
        console.error('Silent token acquisition failed:', error);
        // Fallback to interactive
        try {
            const response = await msalInstance.acquireTokenRedirect(silentRequest);
            accessToken = response.accessToken;
            return accessToken;
        } catch (interactiveError) {
            console.error('Interactive token acquisition failed:', interactiveError);
            return null;
        }
    }
}

// SharePoint API - Create Item
async function createSharePointItem(itemData) {
    const token = await getAccessToken();
    
    const endpoint = `${CONFIG.sharepoint.siteUrl}/_api/web/lists/getbytitle('${CONFIG.sharepoint.listName}')/items`;
    
    const body = {
        __metadata: { type: 'SP.Data.TravelRequestsListItem' },
        Title: `TR-${Date.now()}`,
        RequestedBy: itemData.requestedBy,
        Department: itemData.department,
        TravelType: itemData.travelType,
        EmployeeID: itemData.employeeId,
        TravellerName: itemData.travellerName,
        TravellerAddress: itemData.travellerAddress,
        ContactNumber: itemData.contactNumber,
        TravelDate: itemData.travelDate,
        FromLocationType: itemData.fromLocationType,
        FromAddress: itemData.fromAddress,
        ToLocationType: itemData.toLocationType,
        ToAddress: itemData.toAddress,
        PickupTime: itemData.pickupTime,
        Status: 'Pending',
        SubmittedByEmail: currentUser.username,
        SubmittedByName: currentUser.name
    };
    
    try {
        const response = await fetch(endpoint, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose'
            },
            body: JSON.stringify(body)
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const data = await response.json();
        
        // Upload file if exists
        if (itemData.file) {
            await uploadAttachment(data.d.ID, itemData.file);
        }
        
        return data.d;
    } catch (error) {
        console.error('Error creating SharePoint item:', error);
        throw error;
    }
}

// SharePoint API - Upload Attachment
async function uploadAttachment(itemId, file) {
    const token = await getAccessToken();
    
    const endpoint = `${CONFIG.sharepoint.siteUrl}/_api/web/lists/getbytitle('${CONFIG.sharepoint.listName}')/items(${itemId})/AttachmentFiles/add(FileName='${file.name}')`;
    
    try {
        const arrayBuffer = await file.arrayBuffer();
        
        const response = await fetch(endpoint, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose'
            },
            body: arrayBuffer
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        return await response.json();
    } catch (error) {
        console.error('Error uploading attachment:', error);
        throw error;
    }
}

// SharePoint API - Get Items
async function getSharePointItems(filter = '') {
    const token = await getAccessToken();
    
    let endpoint = `${CONFIG.sharepoint.siteUrl}/_api/web/lists/getbytitle('${CONFIG.sharepoint.listName}')/items?$select=*,AttachmentFiles&$expand=AttachmentFiles&$orderby=Created desc`;
    
    if (filter) {
        endpoint += `&$filter=${filter}`;
    }
    
    try {
        const response = await fetch(endpoint, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Accept': 'application/json;odata=verbose'
            }
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const data = await response.json();
        return data.d.results;
    } catch (error) {
        console.error('Error fetching SharePoint items:', error);
        throw error;
    }
}

// SharePoint API - Update Item
async function updateSharePointItem(itemId, updates) {
    const token = await getAccessToken();
    
    const endpoint = `${CONFIG.sharepoint.siteUrl}/_api/web/lists/getbytitle('${CONFIG.sharepoint.listName}')/items(${itemId})`;
    
    // First, get the item's etag
    const getResponse = await fetch(endpoint, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Accept': 'application/json;odata=verbose'
        }
    });
    
    const itemData = await getResponse.json();
    const etag = itemData.d.__metadata.etag;
    
    const body = {
        __metadata: { type: 'SP.Data.TravelRequestsListItem' },
        ...updates,
        ProcessedBy: currentUser.name,
        ProcessedDate: new Date().toISOString()
    };
    
    try {
        const response = await fetch(endpoint, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
                'IF-MATCH': etag,
                'X-HTTP-Method': 'MERGE'
            },
            body: JSON.stringify(body)
        });
        
        if (!response.ok && response.status !== 204) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        return true;
    } catch (error) {
        console.error('Error updating SharePoint item:', error);
        throw error;
    }
}

// Form submission
document.addEventListener('DOMContentLoaded', function() {
    initializeMsal();
    
    const form = document.getElementById('travelRequestForm');
    if (form) {
        form.addEventListener('submit', async function(e) {
            e.preventDefault();
            
            const submitBtn = form.querySelector('button[type="submit"]');
            submitBtn.disabled = true;
            submitBtn.textContent = 'Submitting...';
            
            try {
                const formData = {
                    requestedBy: document.getElementById('requestedBy').value,
                    department: document.getElementById('department').value,
                    travelType: document.querySelector('input[name="travelType"]:checked').value,
                    employeeId: document.getElementById('employeeId').value,
                    travellerName: document.getElementById('travellerName').value,
                    travellerAddress: document.getElementById('travellerAddress').value,
                    contactNumber: document.getElementById('contactNumber').value,
                    travelDate: document.getElementById('travelDate').value,
                    fromLocationType: document.querySelector('input[name="fromLocation"]:checked').value,
                    fromAddress: document.getElementById('fromAddress').value,
                    toLocationType: document.querySelector('input[name="toLocation"]:checked').value,
                    toAddress: document.getElementById('toAddress').value,
                    pickupTime: document.getElementById('pickupTime').value,
                    file: document.getElementById('ticketFile').files[0]
                };
                
                await createSharePointItem(formData);
                
                showAlert('submitAlert', 'Travel request submitted successfully!', 'success');
                form.reset();
                document.getElementById('fileName').textContent = '';
                
                // Pre-fill user name again
                document.getElementById('requestedBy').value = currentUser.name;
                
            } catch (error) {
                console.error('Submission error:', error);
                showAlert('submitAlert', 'Failed to submit request. Please try again.', 'error');
            } finally {
                submitBtn.disabled = false;
                submitBtn.textContent = 'Submit Request';
            }
        });
    }
    
    // File upload handler
    const fileInput = document.getElementById('ticketFile');
    if (fileInput) {
        fileInput.addEventListener('change', function() {
            const fileName = document.getElementById('fileName');
            if (this.files.length > 0) {
                fileName.textContent = `Selected: ${this.files[0].name}`;
            } else {
                fileName.textContent = '';
            }
        });
    }
});

// Tab switching
function switchTab(tabName) {
    // Hide all tabs
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });
    
    // Remove active from all nav tabs
    document.querySelectorAll('.nav-tab').forEach(tab => {
        tab.classList.remove('active');
    });
    
    // Show selected tab
    document.getElementById(tabName + 'Tab').classList.add('active');
    
    // Set active nav tab
    event.target.classList.add('active');
    
    // Load data for admin tabs
    if (tabName === 'pending') {
        loadPendingRequests();
    } else if (tabName === 'processed') {
        loadProcessedRequests();
    }
}

// Load pending requests
async function loadPendingRequests() {
    const container = document.getElementById('pendingTableContainer');
    container.innerHTML = '<div class="loading"><div class="spinner"></div><p>Loading requests...</p></div>';
    
    try {
        let filter = "Status eq 'Pending'";
        
        const filterDate = document.getElementById('filterDate').value;
        if (filterDate) {
            filter += ` and TravelDate eq '${filterDate}'`;
        }
        
        const items = await getSharePointItems(filter);
        
        if (items.length === 0) {
            container.innerHTML = '<p style="text-align: center; padding: 40px; color: #666;">No pending requests found.</p>';
            return;
        }
        
        let tableHTML = `
            <table>
                <thead>
                    <tr>
                        <th><input type="checkbox" id="selectAllCheckbox" onchange="toggleSelectAll()"></th>
                        <th>Request ID</th>
                        <th>Submitted By</th>
                        <th>Department</th>
                        <th>Traveller</th>
                        <th>Travel Date</th>
                        <th>From</th>
                        <th>To</th>
                        <th>Pickup Time</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        items.forEach(item => {
            tableHTML += `
                <tr>
                    <td><input type="checkbox" class="request-checkbox" value="${item.ID}"></td>
                    <td>${item.Title}</td>
                    <td>${item.SubmittedByName || item.RequestedBy}</td>
                    <td>${item.Department}</td>
                    <td>${item.TravellerName}</td>
                    <td>${new Date(item.TravelDate).toLocaleDateString()}</td>
                    <td>${item.FromLocationType}: ${item.FromAddress.substring(0, 30)}...</td>
                    <td>${item.ToLocationType}: ${item.ToAddress.substring(0, 30)}...</td>
                    <td>${item.PickupTime}</td>
                    <td>
                        <button class="btn btn-danger" style="padding: 5px 10px; font-size: 12px;" onclick="declineRequest(${item.ID})">
                            Decline
                        </button>
                        <button class="btn btn-secondary" style="padding: 5px 10px; font-size: 12px;" onclick="viewDetails(${item.ID})">
                            View
                        </button>
                    </td>
                </tr>
            `;
        });
        
        tableHTML += '</tbody></table>';
        container.innerHTML = tableHTML;
        
    } catch (error) {
        console.error('Error loading pending requests:', error);
        container.innerHTML = '<p style="text-align: center; padding: 40px; color: #dc3545;">Failed to load requests. Please try again.</p>';
    }
}

// Load processed requests
async function loadProcessedRequests() {
    const container = document.getElementById('processedTableContainer');
    container.innerHTML = '<div class="loading"><div class="spinner"></div><p>Loading requests...</p></div>';
    
    try {
        let filter = "(Status eq 'Approved' or Status eq 'Declined')";
        
        const fromDate = document.getElementById('processedFromDate').value;
        const toDate = document.getElementById('processedToDate').value;
        const status = document.getElementById('statusFilter').value;
        
        if (fromDate) {
            filter += ` and ProcessedDate ge '${fromDate}'`;
        }
        if (toDate) {
            filter += ` and ProcessedDate le '${toDate}'`;
        }
        if (status) {
            filter = `Status eq '${status}'` + (fromDate || toDate ? ` and ${filter}` : '');
        }
        
        const items = await getSharePointItems(filter);
        
        if (items.length === 0) {
            container.innerHTML = '<p style="text-align: center; padding: 40px; color: #666;">No processed requests found.</p>';
            return;
        }
        
        let tableHTML = `
            <table>
                <thead>
                    <tr>
                        <th>Request ID</th>
                        <th>Submitted By</th>
                        <th>Department</th>
                        <th>Traveller</th>
                        <th>Travel Date</th>
                        <th>Status</th>
                        <th>Processed By</th>
                        <th>Processed Date</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        items.forEach(item => {
            const statusClass = item.Status === 'Approved' ? 'status-approved' : 'status-declined';
            tableHTML += `
                <tr>
                    <td>${item.Title}</td>
                    <td>${item.SubmittedByName || item.RequestedBy}</td>
                    <td>${item.Department}</td>
                    <td>${item.TravellerName}</td>
                    <td>${new Date(item.TravelDate).toLocaleDateString()}</td>
                    <td><span class="status-badge ${statusClass}">${item.Status}</span></td>
                    <td>${item.ProcessedBy || '-'}</td>
                    <td>${item.ProcessedDate ? new Date(item.ProcessedDate).toLocaleDateString() : '-'}</td>
                    <td>
                        <button class="btn btn-secondary" style="padding: 5px 10px; font-size: 12px;" onclick="viewDetails(${item.ID})">
                            View
                        </button>
                    </td>
                </tr>
            `;
        });
        
        tableHTML += '</tbody></table>';
        container.innerHTML = tableHTML;
        
    } catch (error) {
        console.error('Error loading processed requests:', error);
        container.innerHTML = '<p style="text-align: center; padding: 40px; color: #dc3545;">Failed to load requests. Please try again.</p>';
    }
}

// Toggle select all checkboxes
function toggleSelectAll() {
    const selectAll = document.getElementById('selectAllCheckbox');
    const checkboxes = document.querySelectorAll('.request-checkbox');
    
    checkboxes.forEach(checkbox => {
        checkbox.checked = selectAll.checked;
    });
}

// Clear filter
function clearFilter() {
    document.getElementById('filterDate').value = '';
    loadPendingRequests();
}

// Decline request
async function declineRequest(itemId) {
    const reason = prompt('Please enter the reason for declining this request:');
    
    if (!reason) return;
    
    try {
        await updateSharePointItem(itemId, {
            Status: 'Declined',
            DeclineReason: reason
        });
        
        showAlert('pendingAlert', 'Request declined successfully.', 'success');
        loadPendingRequests();
        
    } catch (error) {
        console.error('Error declining request:', error);
        showAlert('pendingAlert', 'Failed to decline request. Please try again.', 'error');
    }
}

// View details
async function viewDetails(itemId) {
    try {
        const items = await getSharePointItems(`ID eq ${itemId}`);
        const item = items[0];
        
        let detailsHTML = `
            <div style="max-width: 600px; margin: 20px auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
                <h2 style="color: #667eea; margin-bottom: 20px;">Request Details - ${item.Title}</h2>
                
                <div style="margin-bottom: 15px;">
                    <strong>Requested By:</strong> ${item.RequestedBy}
                </div>
                <div style="margin-bottom: 15px;">
                    <strong>Department:</strong> ${item.Department}
                </div>
                <div style="margin-bottom: 15px;">
                    <strong>Travel Type:</strong> ${item.TravelType}
                </div>
                <div style="margin-bottom: 15px;">
                    <strong>Employee ID:</strong> ${item.EmployeeID || 'N/A'}
                </div>
                <div style="margin-bottom: 15px;">
                    <strong>Traveller Name:</strong> ${item.TravellerName}
                </div>
                <div style="margin-bottom: 15px;">
                    <strong>Traveller Address:</strong> ${item.TravellerAddress}
                </div>
                <div style="margin-bottom: 15px;">
                    <strong>Contact Number:</strong> ${item.ContactNumber}
                </div>
                <div style="margin-bottom: 15px;">
                    <strong>Travel Date:</strong> ${new Date(item.TravelDate).toLocaleDateString()}
                </div>
                <div style="margin-bottom: 15px;">
                    <strong>Pickup Time:</strong> ${item.PickupTime}
                </div>
                <div style="margin-bottom: 15px;">
                    <strong>From:</strong> ${item.FromLocationType} - ${item.FromAddress}
                </div>
                <div style="margin-bottom: 15px;">
                    <strong>To:</strong> ${item.ToLocationType} - ${item.ToAddress}
                </div>
                <div style="margin-bottom: 15px;">
                    <strong>Status:</strong> <span class="status-badge status-${item.Status.toLowerCase()}">${item.Status}</span>
                </div>
                
                ${item.AttachmentFiles && item.AttachmentFiles.results.length > 0 ? `
                    <div style="margin-bottom: 15px;">
                        <strong>Attachments:</strong><br>
                        ${item.AttachmentFiles.results.map(file => 
                            `<a href="${file.ServerRelativeUrl}" target="_blank" style="color: #667eea;">${file.FileName}</a>`
                        ).join('<br>')}
                    </div>
                ` : ''}
                
                ${item.DeclineReason ? `
                    <div style="margin-bottom: 15px;">
                        <strong>Decline Reason:</strong> ${item.DeclineReason}
                    </div>
                ` : ''}
                
                <button class="btn btn-secondary" onclick="closeModal()" style="margin-top: 20px;">Close</button>
            </div>
        `;
        
        // Create modal overlay
        const modal = document.createElement('div');
        modal.id = 'detailsModal';
        modal.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.7);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 1000;
            overflow-y: auto;
        `;
        modal.innerHTML = detailsHTML;
        modal.onclick = function(e) {
            if (e.target === modal) closeModal();
        };
        
        document.body.appendChild(modal);
        
    } catch (error) {
        console.error('Error viewing details:', error);
        alert('Failed to load request details.');
    }
}

// Close modal
function closeModal() {
    const modal = document.getElementById('detailsModal');
    if (modal) {
        modal.remove();
    }
}

// Send selected requests to transport
async function sendToTransport() {
    const checkboxes = document.querySelectorAll('.request-checkbox:checked');
    
    if (checkboxes.length === 0) {
        showAlert('pendingAlert', 'Please select at least one request.', 'error');
        return;
    }
    
    const sendBtn = document.getElementById('sendBtn');
    sendBtn.disabled = true;
    sendBtn.textContent = 'Sending...';
    
    try {
        // Get selected items
        const selectedIds = Array.from(checkboxes).map(cb => cb.value);
        const items = await getSharePointItems(`ID in (${selectedIds.join(',')})`);
        
        // Create email draft using Microsoft Graph API
        const emailBody = generateEmailBody(items);
        
        const token = await getAccessToken();
        
        const draft = {
            subject: `${CONFIG.transport.subject}${new Date().toLocaleDateString()}`,
            body: {
                contentType: "HTML",
                content: emailBody
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: CONFIG.transport.email
                    }
                }
            ]
        };
        
        const response = await fetch('https://graph.microsoft.com/v1.0/me/messages', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(draft)
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const draftMessage = await response.json();
        
        // Update items as approved
        for (const id of selectedIds) {
            await updateSharePointItem(id, { Status: 'Approved' });
        }
        
        // Open draft in Outlook
        window.open(`https://outlook.office.com/mail/deeplink/compose/${draftMessage.id}`, '_blank');
        
        showAlert('pendingAlert', 'Draft email created successfully! Please check your Outlook.', 'success');
        
        // Reload pending requests
        setTimeout(() => {
            loadPendingRequests();
        }, 2000);
        
    } catch (error) {
        console.error('Error sending to transport:', error);
        showAlert('pendingAlert', 'Failed to create draft email. Please try again.', 'error');
    } finally {
        sendBtn.disabled = false;
        sendBtn.textContent = '📧 Send Selected to Transport';
    }
}

// Generate email body
function generateEmailBody(items) {
    let body = `
        <html>
        <head>
            <style>
                body { font-family: 'Segoe UI', Arial, sans-serif; }
                table { border-collapse: collapse; width: 100%; margin-top: 20px; }
                th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
                th { background-color: #667eea; color: white; }
                tr:nth-child(even) { background-color: #f2f2f2; }
                h2 { color: #667eea; }
            </style>
        </head>
        <body>
            <h2>Travel Requests for ${new Date().toLocaleDateString()}</h2>
            <p>Dear Transport Team,</p>
            <p>Please arrange the following travel requests:</p>
            
            <table>
                <thead>
                    <tr>
                        <th>Request ID</th>
                        <th>Traveller Name</th>
                        <th>Contact</th>
                        <th>Travel Date</th>
                        <th>Pickup Time</th>
                        <th>From</th>
                        <th>To</th>
                        <th>Department</th>
                    </tr>
                </thead>
                <tbody>
    `;
    
    items.forEach(item => {
        body += `
            <tr>
                <td>${item.Title}</td>
                <td>${item.TravellerName}</td>
                <td>${item.ContactNumber}</td>
                <td>${new Date(item.TravelDate).toLocaleDateString()}</td>
                <td>${item.PickupTime}</td>
                <td>${item.FromLocationType}: ${item.FromAddress}</td>
                <td>${item.ToLocationType}: ${item.ToAddress}</td>
                <td>${item.Department}</td>
            </tr>
        `;
    });
    
    body += `
                </tbody>
            </table>
            
            <p style="margin-top: 30px;">
                <strong>Total Requests:</strong> ${items.length}
            </p>
            
            <p>Thank you,<br>MIT Travel Request System</p>
        </body>
        </html>
    `;
    
    return body;
}

// Show alert message
function showAlert(containerId, message, type) {
    const container = document.getElementById(containerId);
    const alertClass = type === 'success' ? 'alert-success' : 'alert-error';
    
    container.innerHTML = `
        <div class="alert ${alertClass}">
            ${message}
        </div>
    `;
    
    // Auto-hide after 5 seconds
    setTimeout(() => {
        container.innerHTML = '';
    }, 5000);
}
