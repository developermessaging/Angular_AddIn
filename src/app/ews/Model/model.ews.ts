import { GeneralDropDownModel } from "../../Models/generalDropDownModel";
import { resolveTimingValue } from "@angular/animations/browser/src/util";
import { strict } from "assert";
import { Observable } from "rxjs";

export class EWSModel {
    public EWSPayload: string;
    public EWSResponse:string;
    public DisplayEWSResponse:string;
    public itemID: string;


    constructor() {
        this.itemID = Office.context.mailbox.item.itemId;
        this.EWSPayload = this.GetItemSample(this.itemID);  
        debugger;  
    }

private GetItemSample(_itemID:string):string{
    var retval:string = 
`<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" 
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Header>
    <RequestServerVersion Version="Exchange2013" 
      xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />
  </soap:Header>
  <soap:Body>
    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
      <ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Body"/>
          <t:FieldURI FieldURI="message:BccRecipients"/>
        </t:AdditionalProperties>
      </ItemShape>
      <ItemIds>
        <t:ItemId Id="${_itemID}"/>
      </ItemIds>
    </GetItem>
  </soap:Body>
</soap:Envelope>
`;
    

    return retval;
}

}