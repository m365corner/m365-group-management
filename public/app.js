const msalInstance = new msal.PublicClientApplication({
    auth: {
        clientId: "<client-id-goes-here>",
        authority: "https://login.microsoftonline.com/<tenant-id-goes-here>",
        redirectUri: "http://localhost:8000",
    },
});

// Login function
async function login() {
    try {
        const loginResponse = await msalInstance.loginPopup({
            scopes: ["User.ReadWrite.All", "Group.Read.All", "Directory.ReadWrite.All", "Mail.Send"],
        });
        msalInstance.setActiveAccount(loginResponse.account);
        alert("Login successful.");

        // Populate dropdowns after successful login
        await populateDropdowns();
    } catch (error) {
        console.error("Login error:", error);
        alert("Login failed.");
    }
}

// Logout function
function logout() {
    msalInstance.logoutPopup().then(() => alert("Logout successful."));
}

// Populate UPN filter dynamically
async function populateUPNFilter() {
    try {
        const response = await callGraphApi("/users?$select=userPrincipalName");
        const upnFilter = document.getElementById("upnFilter");

        response.value.forEach(user => {
            const option = document.createElement("option");
            option.value = user.userPrincipalName;
            option.textContent = user.userPrincipalName;
            upnFilter.appendChild(option);
        });
    } catch (error) {
        console.error("Error fetching UPNs:", error);
        alert("Failed to populate UserPrincipalName filter.");
    }
}

// Populate Group Name filter dynamically
async function populateGroupNameFilter() {
    try {
        const response = await callGraphApi("/groups?$select=displayName");
        const groupNameFilter = document.getElementById("groupNameFilter");

        response.value.forEach(group => {
            const option = document.createElement("option");
            option.value = group.displayName;
            option.textContent = group.displayName;
            groupNameFilter.appendChild(option);
        });
    } catch (error) {
        console.error("Error fetching Group Names:", error);
        alert("Failed to populate Group Name filter.");
    }
}

// Populate all dropdowns
async function populateDropdowns() {
    await Promise.all([populateUPNFilter(), populateGroupNameFilter()]);
}

// Call Graph API function
async function callGraphApi(endpoint, method = "GET", body = null) {
    const account = msalInstance.getActiveAccount();
    if (!account) {
        throw new Error("No active account. Please login first.");
    }

    try {
        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: ["User.ReadWrite.All", "Directory.ReadWrite.All", "Mail.Send"],
            account: account,
        });

        const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
            method,
            headers: {
                Authorization: `Bearer ${tokenResponse.accessToken}`,
                "Content-Type": "application/json",
            },
            body: body ? JSON.stringify(body) : null,
        });

        if (response.ok) {
            const contentType = response.headers.get("content-type");
            if (contentType && contentType.includes("application/json")) {
                return await response.json(); // Parse JSON response
            }
            return {}; // Return empty object for non-JSON responses like 204 No Content
        } else {
            const errorText = await response.text();
            console.error("Graph API error response:", errorText);
            throw new Error(`Graph API call failed: ${response.status} ${response.statusText}`);
        }
    } catch (error) {
        console.error("Error in callGraphApi:", error.message);
        throw error;
    }
}

// Initialize page
function initializePage() {
    console.log("Page initialized. Please login to populate dropdowns.");
}

// Attach page initialization to DOMContentLoaded
document.addEventListener("DOMContentLoaded", initializePage);




async function searchMembership() {
    const upn = document.getElementById("upnFilter").value;
    const groupName = document.getElementById("groupNameFilter").value;

    if (!upn && !groupName) {
        alert("Please select a UserPrincipalName or Group Name to search.");
        return;
    }

    try {
        let response;

        if (upn) {
            response = await callGraphApi(`/users/${encodeURIComponent(upn)}/memberOf?$select=displayName,groupTypes`);
            /*code added by ME*/ 
            if (!response.value || response.value.length === 0) {
                alert("The selected user is not part of any group.");
                clearTable();
                return;
            }


            populateTable(response.value.map(group => ({
                userPrincipalName: upn,
                group: group.displayName,
                groupType: group.groupTypes && group.groupTypes[0] ? group.groupTypes[0] : "N/A",
            })));
        } else if (groupName) {
            response = await callGraphApi(`/groups?$filter=displayName eq '${groupName}'&$expand=members($select=userPrincipalName,mail)`);
            const group = response.value[0];
              /*code added by ME*/ 
            if (!group || !group.members || group.members.length === 0) {
                alert("The selected group does not have any members.");
                clearTable();
                return;
            }


            if (group && group.members) {
                populateTable(group.members.map(member => ({
                    userPrincipalName: member.userPrincipalName || "N/A",
                    group: group.displayName,
                    groupType: group.groupTypes && group.groupTypes[0] ? group.groupTypes[0] : "N/A",
                    userMail: member.mail || "N/A",
                })));
            } else {
                alert("No users found in the selected group.");
            }
        }
    } catch (error) {
        console.error("Error searching membership:", error);
        alert("Failed to retrieve membership data.");
    }
}

// Populate the result table
function populateTable(data) {
    const outputHeader = document.getElementById("outputHeader");
    const outputBody = document.getElementById("outputBody");

    outputHeader.innerHTML = `
        <th>UserPrincipalName</th>
        <th>Group</th>
        <th>Group Type</th>
        <th>User Mail</th>
    `;

    outputBody.innerHTML = data
        .map(row => `
            <tr>
                <td>${row.userPrincipalName || "N/A"}</td>
                <td>${row.group || "N/A"}</td>
                <td>${row.groupType || "N/A"}</td>
                <td>${row.userMail || "N/A"}</td>
            </tr>
        `)
        .join("");
}

// Send report as mail
// Send Report as Mail
async function sendReportAsMail() {
    const recipientEmail = document.getElementById("recipientEmail").value;

    if (!recipientEmail) {
        alert("Please enter a valid recipient email.");
        return;
    }

    // Extract data from the table
    const tableHeaders = [...document.querySelectorAll("#outputHeader th")].map(th => th.textContent);
    const tableRows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    if (tableRows.length === 0) {
        alert("No data to send. Please retrieve and display user details first.");
        return;
    }

    // Format the email body as an HTML table
    const emailTable = `
        <table border="1" style="border-collapse: collapse; width: 100%;">
            <thead>
                <tr>${tableHeaders.map(header => `<th>${header}</th>`).join("")}</tr>
            </thead>
            <tbody>
                ${tableRows
                    .map(
                        row => `<tr>${row.map(cell => `<td>${cell}</td>`).join("")}</tr>`
                    )
                    .join("")}
            </tbody>
        </table>
    `;

    // Email content
    const email = {
        message: {
            subject: "User Report from M365 User Management Tool",
            body: {
                contentType: "HTML",
                content: `
                    <p>Dear Administrator,</p>
                    <p>Please find below the user report generated by the M365 User Management Tool:</p>
                    ${emailTable}
                    <p>Regards,<br>M365 User Management Team</p>
                `
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: recipientEmail
                    }
                }
            ]
        }
    };

    try {
        const response = await callGraphApi("/me/sendMail", "POST", email);
        alert("Report sent successfully!");
        console.log("Mail Response:", response);
    } catch (error) {
        console.error("Error sending report:", error);
        alert("Failed to send the report. Please try again.");
    }
}


// Download report as CSV
function downloadReportAsCSV() {
    const headers = [...document.querySelectorAll("#outputHeader th")].map(th => th.textContent);
    const rows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    if (rows.length === 0) {
        alert("No data to download.");
        return;
    }

    const csvContent = [headers.join(","), ...rows.map(r => r.join(","))].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "Disabled_Users_Report.csv";
    a.click();
    URL.revokeObjectURL(url);
}


function resetScreen() {
   

    document.getElementById("outputHeader").innerHTML = "";
    document.getElementById("outputBody").innerHTML = "";
    alert("Screen has been reset.");
}


