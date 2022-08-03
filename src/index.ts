import * as common from './common';
import * as nodeApi from 'azure-devops-node-api';
import * as WorkItemTrackingApi from 'azure-devops-node-api/WorkItemTrackingApi';
import * as WorkItemTrackingInterfaces from 'azure-devops-node-api/interfaces/WorkItemTrackingInterfaces';

import * as fs from 'fs';

(async () => {
    // Load the APIs
    const webApi: nodeApi.WebApi = await common.getWebApi();
    const witApi: WorkItemTrackingApi.IWorkItemTrackingApi = await webApi.getWorkItemTrackingApi();

    // Query the project to retreive the Id field for all workitems
    const queryResult = await witApi.queryByWiql({ query: 'select [System.Id] From WorkItems' });

    // Map this into an array of number
    const workItemIds = queryResult.workItems.map(item => item.id);

    // Retrive fully-expanded workitems. Note this is limited to 200 records, so
    // pagniation control would be required for a bigger project
    const workItems = await witApi.getWorkItems(workItemIds, undefined, undefined, WorkItemTrackingInterfaces.WorkItemExpand.All);

    // Write the JSON to file
    fs.writeFileSync('./workItems.json',JSON.stringify(workItems));
})()