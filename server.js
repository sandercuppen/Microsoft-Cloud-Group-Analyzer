const express = require('express');
const path = require('path');
const helper = require('./helper');

const app = express();
const port = process.env.PORT || 3001;

// Serve static files from the public directory
app.use(express.static('public'));

// API endpoints
app.get('/api/auth-status', async (req, res) => {
    try {
        const token = await helper.getToken();
        const user = await helper.callApi('https://graph.microsoft.com/v1.0/me', token.accessToken);
        res.json({ 
            authenticated: true, 
            userPrincipalName: user.userPrincipalName 
        });
    } catch (error) {
        res.json({ authenticated: false });
    }
});

app.get('/api/analyze', async (req, res) => {
    try {
        const token = await helper.getToken();
        const groupId = req.query.groupId || 'all';

        let groups = [];
        if (groupId === 'all') {
            // Get all groups
            groups = await helper.getAllWithNextLink(token.accessToken, 'https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,groupTypes,membershipRule,resourceProvisioningOptions,securityEnabled,visibility,mailEnabled,mailNickname,membershipRuleProcessingState');
        } else {
            // Get specific group
            const group = await helper.callApi(`https://graph.microsoft.com/v1.0/groups/${groupId}?$select=id,displayName,description,groupTypes,membershipRule,resourceProvisioningOptions,securityEnabled,visibility,mailEnabled,mailNickname,membershipRuleProcessingState`, token.accessToken);
            if (group) {
                groups = [group];
            }
        }

        // For each group, get members and owners count
        for (let group of groups) {
            let members = await helper.getAllWithNextLink(token.accessToken, `https://graph.microsoft.com/v1.0/groups/${group.id}/members?$select=id`);
            let owners = await helper.getAllWithNextLink(token.accessToken, `https://graph.microsoft.com/v1.0/groups/${group.id}/owners?$select=id`);
            
            group.memberCount = members ? members.length : 0;
            group.ownerCount = owners ? owners.length : 0;
            
            // Add any potential issues
            group.issues = [];
            if (group.ownerCount === 0) {
                group.issues.push('No owners assigned');
            }
            if (group.memberCount === 0) {
                group.issues.push('No members in group');
            }
        }

        res.json(groups);
    } catch (error) {
        console.error('Error analyzing groups:', error);
        res.status(500).json({ error: 'Analysis failed' });
    }
});

app.get('/api/export', async (req, res) => {
    try {
        const token = await helper.getToken();
        const groupId = req.query.groupId || 'all';
        const format = req.query.format;

        // Reuse the analyze endpoint logic
        let groups = [];
        if (groupId === 'all') {
            groups = await helper.getAllWithNextLink(token.accessToken, 'https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,groupTypes,membershipRule,resourceProvisioningOptions,securityEnabled,visibility,mailEnabled,mailNickname,membershipRuleProcessingState');
        } else {
            const group = await helper.callApi(`https://graph.microsoft.com/v1.0/groups/${groupId}`, token.accessToken);
            if (group) {
                groups = [group];
            }
        }

        if (format === 'csv') {
            const csv = await helper.exportCSV(groups, null);
            res.setHeader('Content-Type', 'text/csv');
            res.setHeader('Content-Disposition', 'attachment; filename=group-analysis.csv');
            res.send(csv);
        } else if (format === 'json') {
            res.setHeader('Content-Type', 'application/json');
            res.setHeader('Content-Disposition', 'attachment; filename=group-analysis.json');
            res.json(groups);
        } else {
            res.status(400).json({ error: 'Invalid format' });
        }
    } catch (error) {
        res.status(500).json({ error: 'Export failed' });
    }
});

// Serve the main page for all other routes
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Start the server
if (require.main === module) {
    app.listen(port, () => {
        console.log(`Web interface running at http://localhost:${port}`);
    });
}
