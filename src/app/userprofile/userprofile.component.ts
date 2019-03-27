import { Component } from '@angular/core';

const template = require('./userprofile.component.html');

@Component({
    selector: 'userprofile',
    template
})
export default class userprofile {

    private _UserProfile: Office.UserProfile = Office.context.mailbox.userProfile;
    public get UserProfile(): Office.UserProfile {
        return this._UserProfile;
    }
    public set UserProfile(value: Office.UserProfile) {
        this._UserProfile = value;

    }

    constructor() {
        this._UserProfile = Office.context.mailbox.userProfile;
        
    }

    //#region
    splitArray(data:any):string
    {
        var retval:string="";

        data.array.forEach(element => {
            retval += ( element + ";" );
        });


        return retval;
    }
    //#endregion

}