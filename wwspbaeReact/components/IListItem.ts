import { PersonaBase } from "office-ui-fabric-react/lib/components/Persona/Persona.base";

export interface IListItem {
    NotifyToId: any[];  
    Title?: string;  
    Location: string;
    Attendees: string;
    OtherCadenceAttendeesId: any [];
    Status: string;
    AccountName: string;
    ExecutiveSummary: string;
    DetailedSummary: string;
    DateofVisit: string;
    Purposeofthemeeting: any [];
    SubmitterId: number;
    AuthorId: number;
    EditorId: number;
}  