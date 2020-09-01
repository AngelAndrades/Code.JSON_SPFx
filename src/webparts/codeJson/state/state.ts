export type StateKey = 'pageSize' | 'dsAppendData' | 'dsImportData' | 'organization' | 'contactName' | 'contactEmail' | 'vcs' | 'homeLink' | 'importList' | 'appendList' | 'spLink';

export interface State {
    pageSize: number;
    dsAppendData: [];
    dsImportData: [];
    organization: string;
    contactName: string;
    contactEmail: string;
    vcs: string;
    homeLink: string;
    importList: string;
    appendList: string;
    spLink: string;
}

export const INITIAL_STATE: State = {
    pageSize: 5,
    dsAppendData: [],
    dsImportData: [],
    organization: 'OIT EPMO',
    contactName: 'EPMO Code Sharing Services',
    contactEmail: 'OSSOFT@va.gov',
    vcs: 'git',
    homeLink: 'https://github.com/department-of-veterans-affairs',
    importList: null,
    appendList: null,
    spLink: null,
};