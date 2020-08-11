export type StateKey = 'pageSize' | 'dsAppendData' | 'dsImportData' | 'organization' | 'contactName' | 'contactEmail';

export interface State {
    pageSize: number;
    dsAppendData: [];
    dsImportData: [];
    organization: string;
    contactName: string;
    contactEmail: string;
}

export const INITIAL_STATE: State = {
    pageSize: 5,
    dsAppendData: [],
    dsImportData: [],
    organization: 'OIT EPMO',
    contactName: 'EPMO Code Sharing Services',
    contactEmail: 'OSSOFT@va.gov'
};