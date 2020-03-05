/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = () => {
    console.log("Initialize function called insertDefaultCard");
    //console.log(Office.context.roamingSettings.get("RoamingSettingsTest"));
  };

  // Add any ui-less function here
  // Helper function to add a status message to the info bar.


 async function insertDefaultCard(event) {
  //console.log("insertDefaultCard -try");
  Office.context.mailbox.item.notificationMessages.addAsync("error", {
         type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
         message: "RoamingSettingTest : "
       });
  event.completed();

  // try {
  //   console.log("insertDefaultCard -try");
  //   var roamingsetting = Office.context.roamingSettings.get("RoamingSettingTest");
  //   Office.context.mailbox.item.notificationMessages.addAsync("error", {
  //     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
  //     message: "RoamingSettingTest : " + roamingsetting,
  //   });

  // } catch (error) {
  //   Office.context.mailbox.item.notificationMessages.addAsync("error", {
  //     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
  //     message: "RoamingSettingTest Error : " + JSON.stringify(error)
  //   });
    
  // }
  // finally{
  //   event.completed();
  // }  

  

}




})();
