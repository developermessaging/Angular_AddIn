import { Component, ViewChild } from '@angular/core';
import { EWSModel } from './Model/model.ews';
import beautify from 'xml-beautifier';


const template = require('./ews.component.html');

@Component({
    selector: 'ews',
    template
})
export default class ews {

    @ViewChild('editor') editor;

    private _CurrentItem: Office.Item = Office.context.mailbox.item;
    public get CurrentItem(): Office.Item {
        return this._CurrentItem;
    }
    public set CurrentItem(value: Office.Item) {
        this._CurrentItem = value;

    }

    public Model: EWSModel = new EWSModel();

    public text:string;

    constructor() {
        this.CurrentItem = Office.context.mailbox.item;
        this.Model = new EWSModel();
        debugger;

    }

    ngAfterViewInit() {
        this.editor.setTheme("eclipse");
 
        this.editor.getEditor().setOptions({
            enableBasicAutocompletion: true
        });
        this.editor.config.set('basePath', '/libs/ace');

        this.editor.getEditor().commands.addCommand({
            name: "showOtherCompletions",
            bindKey: "Ctrl-.",
            exec: function (editor) {
 
            }
        });

        this.editor.getSession().setMode("ace/mode/xml");
    }

    async CallEWS() {
        var that: this = this;
        Office.context.mailbox.makeEwsRequestAsync(this.Model.EWSPayload, function (asyncResult) {
            var result = asyncResult.value;
            // var context = asyncResult.context;

            debugger;
            // Process the returned response here.
            that.Model.EWSResponse = result;
            that.Model.DisplayEWSResponse = beautify(that.Model.EWSResponse);
            that.text = that.Model.DisplayEWSResponse;
        });
    }


    public callback(asyncResult: any) {
        var result = asyncResult.value;
        var context = asyncResult.context;
        var that = this;

        //debugger;
        // Process the returned response here.
        that.Model.EWSResponse = result;
        that.Model.DisplayEWSResponse = beautify(that.Model.EWSResponse);

    }
}