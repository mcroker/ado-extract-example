import * as common from './common';
import * as nodeApi from 'azure-devops-node-api';
import * as WorkItemTrackingApi from 'azure-devops-node-api/WorkItemTrackingApi';
import * as WorkItemTrackingInterfaces from 'azure-devops-node-api/interfaces/WorkItemTrackingInterfaces';

import * as fs from 'fs';

(async () => {
    const webApi: nodeApi.WebApi = await common.getWebApi();
    const witApi: WorkItemTrackingApi.IWorkItemTrackingApi = await webApi.getWorkItemTrackingApi();

    const queryResult = await witApi.queryByWiql({ query: 'select [System.Id] From WorkItems' });
    const workItemIds = queryResult.workItems.map(item => item.id);
    const workItems = await witApi.getWorkItems(workItemIds, undefined, undefined, WorkItemTrackingInterfaces.WorkItemExpand.All);
    fs.writeFileSync('./workItems.json',JSON.stringify(workItems));
})()