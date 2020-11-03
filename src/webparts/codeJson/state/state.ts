export type StateKey = 'pageSize' | 'dsAppendData' | 'dsImportData' | 'organization' | 'contactName' | 'contactEmail' | 'vcs' | 'homeLink' | 'importList' | 'appendList' | 'spLink' | 'licensing' | 'disclaimer';

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
    licensing: boolean;
    defaultLicense: string;
    defaultLicenseUrl: string;
    disclaimer: string;
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
    licensing: false,
    defaultLicense: 'Creative Commons Zero v1.0 Universal',
    defaultLicenseUrl: 'https://creativecommons.org/publicdomain/zero/1.0/',
    disclaimer: 'License information not currently available. Information forthcoming.'
};