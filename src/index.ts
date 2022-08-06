import * as common from './common';
import * as nodeApi from 'azure-devops-node-api';
import * as WorkItemTrackingApi from 'azure-devops-node-api/WorkItemTrackingApi';
import * as WorkItemTrackingInterfaces from 'azure-devops-node-api/interfaces/WorkItemTrackingInterfaces';

const MSEC_PER_DAY = 24 * 60 * 60 * 1000;

export enum State {
    new = 'new',
    elborated = 'elborated',
    fdApproved = 'fdApproved',
    inTD = 'inTD',
    tdComplete = 'tdComplete',
    inDev = 'inDev',
    inCodeReview = 'inCodeReview',
    readyForTest = 'readyForTest',
    inTest = 'inTest',
    closed = 'closed',
    removed = 'removed',
    blocked = 'blocked'
}
function toState(x: string): State {
    switch (x) {
        case 'New': return State.new;
        case 'Elaborated': return State.elborated;
        case 'Functional Design Approved': return State.fdApproved;
        case 'In Technical Design': return State.inTD;
        case 'Technical Design Completed': return State.tdComplete;
        case 'In Development': return State.inDev;
        case 'In Code Review': return State.inCodeReview;
        case 'Ready for Test': return State.readyForTest;
        case 'In Test': return State.inTest;
        case 'Closed': return State.closed;
        case 'Removed': return State.removed;
        case 'Blocked': return State.blocked;
        default: console.error('Not found :', x); return undefined;
    }
}

export interface StateSummary {
    count: number,
    storyPoints: number
}

export type StatesSummary = { [key: string]: StateSummary };

export interface WorkItem {
    id: number;
    rev: number;
    state: State;
    title: string;
    url: string;
    storyPoints: number; // Microsoft.VSTS.Scheduling.StoryPoints
    boardColumn: string; // "System.BoardColumn": "New",
    boardColumnDone: boolean; // "System.BoardColumnDone": false,
    StateChangeDate: string; // "Microsoft.VSTS.Common.StateChangeDate": "2021-11-18T14:34:06.54Z",
    areaPath: string; // "System.AreaPath": "BDC\\OAS (roll up)\\Business",
    teamProject: string // "System.TeamProject": "BDC",
    nodeName: string; // "System.NodeName": "Business",
    areaLevel1: string; // "System.AreaLevel1": "BDC",
    areaLevel2: string; // "System.AreaLevel2": "OAS (roll up)",
    areaLevel3: string; // "System.AreaLevel3": "Business",
    authorizedDate: string; // "System.AuthorizedDate": "2022-07-15T16:10:36.237Z",
    revisedDate: string; // "System.RevisedDate": "9999-01-01T00:00:00Z",
    iterationpath: string; // "System.IterationPath": "BDC",
    iterationLevel1: string; // "System.IterationLevel1": "BDC",
    workItemType: string; // "System.WorkItemType": "Hybrid Story",
    reason: string; // "System.Reason": "New",
    createdDate: string; // "System.CreatedDate": "2021-11-18T14:34:06.54Z",
    changedDate: string; // "System.ChangedDate": "2022-07-15T16:10:36.237Z",
    businessBoardColumn: string; // "WEF_76EA049BBA4140FEA4D87B5A9F33458C_Kanban.Column": "01-Work in Progress",
    businessColumnDone: boolean; //  "WEF_76EA049BBA4140FEA4D87B5A9F33458C_Kanban.Column.Done": false,
    businessLane: string; // "WEF_76EA049BBA4140FEA4D87B5A9F33458C_Kanban.Lane": "Pod 5 – Interfaces (Lead: )",
    devBoardColumn: string; // "WEF_B16E2796978A433587ED3C652FE9C636_Kanban.Column": "In Technical Design",
    devColumnDone: boolean; // "WEF_B16E2796978A433587ED3C652FE9C636_Kanban.Column.Done": false,

}

(async () => {
    // Load the APIs
    const webApi: nodeApi.WebApi = await common.getWebApi();
    const witApi: WorkItemTrackingApi.IWorkItemTrackingApi = await webApi.getWorkItemTrackingApi();
    const workitems = await getWorkItemsForDates(witApi, statusDates(5, new Date('2022-07-01')));
    console.log(workitems.map(wiSet => sumarizeWorkItemSet(wiSet)));
})()

function sumarizeWorkItemSet(workitems: WorkItem[]): StatesSummary {
    return workitems.reduce((results, item) => {
        const result: StateSummary = results[item.state] || { count: 0, storyPoints: 0 };
        result.count = result.count += 1;
        result.storyPoints = result.storyPoints += isNaN(item.storyPoints) ? 0 : item.storyPoints;
        results[item.state] = result;
        console.log(results);
        return results;
    }, {} as StatesSummary);
}

function getWorkItemsForDates(witApi: WorkItemTrackingApi.IWorkItemTrackingApi, dates: Date[]): Promise<WorkItem[][]> {
    return Promise.all(dates.map(dt => getWorkItems(witApi, new Date(dt))));
}

async function getWorkItems(witApi: WorkItemTrackingApi.IWorkItemTrackingApi, asOf?: Date): Promise<WorkItem[]> {
    // Query the project to retreive the Id field for all workitems
    const ASOF = (asOf) ? ` ASOF '${strDate(asOf)}'` : '';
    const queryResult = await witApi.queryByWiql({ query: `select [System.Id] From WorkItems where [System.workItemType] = 'Hybrid Story' ${ASOF}` });
    const workItemIds = queryResult.workItems.map(item => item.id);

    // Map this into an array of number
    const batchedWorkItems: number[][] = [];
    const chunkSize = 200;
    for (let i = 0; i < workItemIds.length; i += chunkSize) {
        const chunk = workItemIds.slice(i, i + chunkSize);
        batchedWorkItems.push(chunk);
    }

    const workItems: WorkItem[] = [];
    await Promise.all(batchedWorkItems.map(async ids => {
        workItems.push(...await getWorkItemsBatch(witApi, ids, asOf));
    }));
    return workItems;
}

async function getWorkItemsBatch(witApi: WorkItemTrackingApi.IWorkItemTrackingApi, ids: number[], asOf?: Date): Promise<WorkItem[]> {
    const items = await witApi.getWorkItemsBatch({
        $expand: WorkItemTrackingInterfaces.WorkItemExpand.All,
        errorPolicy: WorkItemTrackingInterfaces.WorkItemErrorPolicy.Omit,
        asOf,
        ids
    });
    return items.filter(item => item).map(item => ({
        id: item.id,
        rev: item.rev,
        state: toState(item.fields['System.State']),
        title: item.fields['System.Title'],
        url: item.url,
        storyPoints: item.fields['Microsoft.VSTS.Scheduling.StoryPoints'],
        boardColumn: item.fields['System.BoardColumn'],
        boardColumnDone: item.fields['System.BoardColumnDone'],
        StateChangeDate: item.fields['Microsoft.VSTS.Common.StateChangeDate'],
        areaPath: item.fields['System.AreaPath'],
        teamProject: item.fields['System.TeamProject'],
        nodeName: item.fields['System.NodeName'],
        areaLevel1: item.fields['System.AreaLevel1'],
        areaLevel2: item.fields['System.AreaLevel2'],
        areaLevel3: item.fields['System.AreaLevel3'],
        authorizedDate: item.fields['System.AuthorizedDate'],
        revisedDate: item.fields['System.RevisedDate'],
        iterationpath: item.fields['System.IterationPath'],
        iterationLevel1: item.fields['System.IterationLevel1'],
        workItemType: item.fields['System.WorkItemType'],
        reason: item.fields['System.Reason'],
        createdDate: item.fields['System.CreatedDate'],
        changedDate: item.fields['System.ChangedDate'],
        businessBoardColumn: item.fields['WEF_76EA049BBA4140FEA4D87B5A9F33458C_Kanban.Column'],
        businessColumnDone: item.fields['WEF_76EA049BBA4140FEA4D87B5A9F33458C_Kanban.Column.Done'],
        businessLane: item.fields['WEF_76EA049BBA4140FEA4D87B5A9F33458C_Kanban.Lane'],
        devBoardColumn: item.fields['WEF_B16E2796978A433587ED3C652FE9C636_Kanban.Column'],
        devColumnDone: item.fields['WEF_B16E2796978A433587ED3C652FE9C636_Kanban.Column.Done']
    }))
}

function statusDates(day: number, first: Date): Date[] {  // Default to Friday
    const tsToday: number = MSEC_PER_DAY * Math.floor(((new Date()).getTime() / MSEC_PER_DAY));
    const today = new Date(tsToday);
    const tsAdjust = (today.getDay() === day)
        ? (MSEC_PER_DAY * (today.getDay() - day + 6)) + 1
        : (MSEC_PER_DAY * (today.getDay() - day - 1)) + 1;
    const tsLast = tsToday - tsAdjust;
    const dts: Date[] = [];
    for (let i = tsLast; i >= first.getTime(); i -= (7 * MSEC_PER_DAY)) {
        dts.push(new Date(i));
    }
    return dts;
}

function strDate(dt: Date | string) {
    const d: Date = (typeof dt === 'string') ? new Date(dt) : dt;
    return d.toISOString().split('T')[0];
}
