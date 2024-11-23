/* global document, Office, module, require, HTMLElement */

Office.onReady((info) => {
  if (!info.host) {
    console.log('Warning: Office host information is not set.', info)
    return
  }

  const elmOfficeJsRequired = document.getElementById('officejs-required-info')
  elmOfficeJsRequired.style = 'display: none !important'

  const elmOutlookApp = document.getElementById('outlook-app')
  elmOutlookApp.innerText = `App connected to ${info.host} on ${info.platform}.`
});
