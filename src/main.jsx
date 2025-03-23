import './sentry-init'
import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import Taskpane from './taskpane.jsx'

/* global document, Office, Sentry */

const container = document.getElementById('root');
const root = createRoot(container);

const spanInit = Sentry.startInactiveSpan({ name: "office-initialize", op: 'officejs.onready' })

Office.onReady((info) => {
  spanInit.setAttributes({ host: info.host, platform: info.platform })
  spanInit.end()

  root?.render(
    <StrictMode>
      <Taskpane hostInfo={info} />
    </StrictMode>
  )
})
