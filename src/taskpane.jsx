
import * as Sentry from "@sentry/react";

const Taskpane = (props) => {
  const { host, platform } = props.hostInfo

  /*
  <p id="officejs-required-info">
    Add-in is not connected to Office application,
    or <code>Office.js</code> failed to load.
  </p>
  */

  return (
    <div>
      <h1>Acme Add-in</h1>
      <p className="info">
        Office web based add-in.
      </p>
      <div id="outlook-app">
        App connected to {host} on {platform}.
      </div>
      <p>
        <button type="button" onClick={() => { throw new Error("Sentry Test Error") }}>Break the world</button>
      </p>
    </div>
  )
}

export default Sentry.withProfiler(Taskpane)
