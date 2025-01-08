const helper = require('../helper');

async function init(accessToken, accessTokenAzure, groupID, groupName, tenantID) {
    let subscriptions = await helper.callApi(`https://management.azure.com/subscriptions?api-version=2020-01-01`, accessTokenAzure)

    // if cannot find subscriptions or has no subsriptions
    if (subscriptions == undefined) {
        forbiddenErrors.push(`Error fetching subscriptions`)
        return null
    } else if (subscriptions?.value?.length == 0) {
        forbiddenErrors.push(`No Azure subscriptions found`)
        return null
    }

    // Get management groups
    let mgGroups = await helper.callApi(`https://management.azure.com/providers/Microsoft.Management/managementGroups?api-version=2020-05-01`, accessTokenAzure)
    const mgLookup = {}
    if (mgGroups?.value) {
        mgGroups.value.forEach(mg => {
            mgLookup[mg.name] = mg.properties?.displayName || mg.name
        })
    }

    // Get role definitions for lookup
    let roleDefinitions = await helper.callApi(`https://management.azure.com/providers/Microsoft.Authorization/roleDefinitions?api-version=2022-04-01`, accessTokenAzure)
    const roleLookup = {}
    if (roleDefinitions?.value) {
        roleDefinitions.value.forEach(role => {
            roleLookup[role.name] = role.properties?.roleName || role.name
        })
    }

    // Use a Set to track unique assignments
    const uniqueAssignments = new Set()
    let array = []
    let counter = 1

    var promise = new Promise((resolve, reject) => {
        subscriptions?.value?.forEach(async subscription => {
            let roleAssignments = await helper.callApi(`https://management.azure.com/subscriptions/${subscription.subscriptionId}/providers/Microsoft.Authorization/roleAssignments?api-version=2022-04-01`, accessTokenAzure)
            
            if (roleAssignments) {
                roleAssignments = roleAssignments?.value?.filter(roleAssignment => roleAssignment?.properties?.principalId == groupID).forEach(res => {
                    if (res.properties.scope.length > 1) { // don't show if scope is empty. If scope is empty, it is the RoleAssignment object. No need to show that again
                        const scope = res?.properties?.scope
                        let name = scope.substring(scope.lastIndexOf('/') + 1)
                        
                        // If this is a management group, use its display name
                        if (scope.includes('/providers/Microsoft.Management/managementGroups/')) {
                            const mgId = scope.split('/managementGroups/')[1].split('/')[0]
                            name = mgLookup[mgId] || mgId
                        }
                        
                        // Get role name from lookup
                        const roleDefId = res.properties.roleDefinitionId
                        const roleId = roleDefId.substring(roleDefId.lastIndexOf('/') + 1)
                        const roleName = roleLookup[roleId] || roleId

                        // Create a unique key for this assignment
                        const assignmentKey = `${scope}|${roleId}|${groupID}`
                        
                        // Only add if we haven't seen this exact assignment before
                        if (!uniqueAssignments.has(assignmentKey)) {
                            uniqueAssignments.add(assignmentKey)
                            array.push({
                                "file": 'azureResource',
                                "groupID": groupID,
                                "groupName": groupName,
                                "service": "Azure Resource",
                                "resourceID": res.id,
                                "name": name,
                                "detailsGroup": `${(groupName.includes('@')) ? 'user' : 'group'} '${groupName}'`,
                                "details": `Role: ${roleName}`
                            })
                        }
                    }
                })
            }

            if (subscriptions?.value?.length == counter) resolve()
            counter++
        });
    })

    return promise.then(() => {
        // Sort the array by scope and role name for better readability
        return array.sort((a, b) => {
            // First sort by name (scope)
            const nameCompare = a.name.localeCompare(b.name)
            if (nameCompare !== 0) return nameCompare
            // Then by role name
            return a.details.localeCompare(b.details)
        })
    })
}

module.exports = { init }