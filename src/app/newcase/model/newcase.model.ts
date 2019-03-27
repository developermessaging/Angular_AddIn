import { GeneralDropDownModel } from "../../Models/generalDropDownModel";
import { resolveTimingValue } from "@angular/animations/browser/src/util";
import { strict } from "assert";
import { Observable } from "rxjs";

export class NewCaseModel {
    CaseNumber: string;
    CaseTitle: string;

    CaseType: GeneralDropDownModel[];
    CaseTypeSelectedIndex: Number;

    CaseSeverity: GeneralDropDownModel[];
    CaseSeveritySelectedIndex: Number;

    CaseSource: GeneralDropDownModel[];
    CaseSourceSelectedIndex: Number;

    SLA: string ;
    SLAVisible: Boolean = true;


    ManualAnnouncementOverride: Boolean;

    private _TextAnnouncement: string;
    public get TextAnnouncement(): string {
        if (this.ManualAnnouncementOverride)
            return this._TextAnnouncement;
        else
            return this.GenerateAnnouncement();
    }
    public set TextAnnouncement(v: string) {
        this._TextAnnouncement = v;
    }


    constructor() {
        this.CaseSeverity = [
            { id: 1, name: "Severity A - HIGH IMPACT" },
            { id: 2, name: "Severity B - MEDIUM IMPACT" },
            { id: 3, name: "Severity C - LOW IMPACT" }
        ];
        this.CaseSeveritySelectedIndex = 3;

        this.CaseType = [
            { id: 1, name: "Case" },
            { id: 2, name: "Problem" },
            { id: 3, name: "Request Assistance" }
        ];
        this.CaseTypeSelectedIndex = 1;

        this.CaseSource = [
            { id: 1, name: "MSSolve" },
            { id: 2, name: "RAVE" }
        ];
        this.CaseSourceSelectedIndex = 1;
        this.ManualAnnouncementOverride = false;

        this._TextAnnouncement = this.GenerateAnnouncement();
    }

    private GenerateAnnouncement(): string {
        var retval: string = 
        `New Case Severity ${ this.CaseSeveritySelectedIndex}
@EMEA Team
Number: ${this.CaseNumber} - ${this.CaseTitle}
SLA: ${ this.SLA }
    `;

        return retval;
    }


}