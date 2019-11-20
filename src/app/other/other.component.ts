import { Component } from '@angular/core';
import { GeneralDropDownModel } from '../Models/generalDropDownModel';
import context from '../context/context.component';

const template = require('./other.component.html');

@Component({
    selector: 'other',
    styleUrls: ["./src/app/other/other.component.css"],
    template
})
export default class other {

    public CurrentItem: Office.Item = Office.context.mailbox.item;
    public Mailbox: Office.Mailbox = Office.context.mailbox;


    public GetRestType: GeneralDropDownModel[] = [
        { id: 1, name: "Beta" },
        { id: 2, name: "V2.0" },
        { id: 3, name: "V1.0" }
    ]
    public SelectedGetRestTypeIndex: number;

    public GetRestUrl(): string {
        return Office.context.mailbox.restUrl;
    }
    public GetewsUrl(): string {
        return Office.context.mailbox.ewsUrl;
    }

    public RestId: string = "";
    public GetRestId(): string {
        if(Office.context.mailbox.item.itemId == null) return "";
        var restType: Office.MailboxEnums.RestVersion;
        switch (this.SelectedGetRestTypeIndex.toString()) {
            case "1":
                restType = Office.MailboxEnums.RestVersion.Beta;
                break;
            case "2":
                restType = Office.MailboxEnums.RestVersion.v2_0;
                break;
            case "3":
                restType = Office.MailboxEnums.RestVersion.v1_0;
                break;

            default:
                restType = Office.MailboxEnums.RestVersion.v1_0;
                break;
        }
        this.RestId = Office.context.mailbox.convertToRestId(Office.context.mailbox.item.itemId, restType);
        return this.RestId;
    }


    constructor() {
        this.SelectedGetRestTypeIndex = 1;
        this.GetRestId(); //refresh internal RestId Variable for first time execution
    }


    async displayNewAppointment() {
        var bodyInvite = require("./other.component.newappointment.html");
        var start = new Date();
        var end = new Date();
        end.setHours(start.getHours() + 1);

        Office.context.mailbox.displayNewAppointmentForm(
            {
                requiredAttendees: ['EmeaMessagingDev@contoso.com'],
                optionalAttendees: ['BroaderAudience@contoso.com'],
                start: start,
                end: end,
                location: 'EMEA Messaging Dev Meeting Room',
                resources: ['projector@contoso.com'],
                subject: 'Sample meeting by Demo Addin in ',
                body: bodyInvite //'Hello World!'
            });

    }


    async displayMessageForm() {
        var itemId: string = Office.context.mailbox.item.itemId;

        Office.context.mailbox.displayMessageForm(itemId);
        Office.context.mailbox.convertToRestId(Office.context.mailbox.item.itemId, Office.MailboxEnums.RestVersion.Beta);

    }


    async displayNewMessageForm() {
        Office.context.mailbox.displayNewMessageForm(
            {
                // Copy the To line from current item.
                toRecipients: Office.context.mailbox.item.to,
                ccRecipients: ['sam@contoso.com'],
                subject: 'Outlook add-ins are cool!',
                htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
                attachments: [
                    {
                        type: 'file',
                        name: 'image.png',
                        url: 'http://contoso.com/image.png',
                        isInline: true
                    }
                ]
            });
    }


    async displayDialogAsync() {
        debugger;
        Office.context.ui.displayDialogAsync("https://forms.office.com/Pages/ResponsePage.aspx?id=26HulZLRrESMASNUupPx8VjUibkIkxFEsoqd86bMiAxUQjhYRjBOOEVNN0FMWTdST1dOOUhFUEpQUC4u&embed=true",
            { height: 60, width: 40, promptBeforeOpen: false }, 
            function (result) {
                debugger;
                var _dlg = result.value; 
                //_dlg.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
            });

    }

    async SetAttendees(){
        var x:string = "0";
        Office.context.mailbox.item.requiredAttendees.setAsync(
            ['CosmosRoom@devmsgpt.onmicrosoft.com', 'MercuryRoom@devmsgpt.onmicrosoft.com'] , this.callbackAttendees );
        debugger;
    }

    async callbackAttendees(){
        var x:string = "0";
        debugger;
    }


    public onChange(event): void {  // event will give you full breif of action
        const newVal = event.target.value;
        var s: string = this.GetRestId();
        //debugger;
        console.log(s);
    }

    async SetFormattedBody() {
        debugger;
        var formattedBody:string = "<img src='https://upload.wikimedia.org/wikipedia/commons/thumb/9/91/Spain_traffic_signal_r100.svg/1024px-Spain_traffic_signal_r100.svg.png'/>";
                    Office.context.mailbox.item.body.setAsync( formattedBody,
                        { coercionType: Office.CoercionType.Html, asyncContext: null });
        }
    

    // async createRange(number){
    //     var items: number[] = [];
    //     for(var i = 1; i <= number; i+= 0.1){
    //        items.push(i);
    //     }
    //     return items;
    //   }

    public copyMessage(val: string) {
        let selBox = document.createElement('textarea');
        selBox.style.position = 'fixed';
        selBox.style.left = '0';
        selBox.style.top = '0';
        selBox.style.opacity = '0';
        selBox.value = val;
        document.body.appendChild(selBox);
        selBox.focus();
        selBox.select();
        document.execCommand('copy');
        document.body.removeChild(selBox);
    }

} 