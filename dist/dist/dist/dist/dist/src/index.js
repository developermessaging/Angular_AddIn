import 'zone.js'; // Required for Angular
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import AppModule from './app/app.module';
Office.initialize = function (reason) {
    document.getElementById('sideload-msg').style.display = 'none';
    // Bootstrap the app
    platformBrowserDynamic().bootstrapModule(AppModule).catch(function (error) { return console.error(error); });
};
//# sourceMappingURL=index.js.map
//# sourceMappingURL=index.js.map
//# sourceMappingURL=index.js.map
//# sourceMappingURL=index.js.map
//# sourceMappingURL=index.js.map