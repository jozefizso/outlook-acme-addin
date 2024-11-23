
import * as Sentry from "@sentry/react";

const Taskpane = (props) => {
  const { host, platform } = props.hostInfo

  var elmMatches = <p>No matches found</p>

  var allMatches = Office.context.mailbox.item.getRegExMatches()
  if (allMatches) {
    elmMatches = (
      <div>
        <p>Your e-mail contains links to these services:</p>
        <p>{JSON.stringify(allMatches.AcmeUrl)}</p>
      </div>
    )
  }

  return (
    <div>
      <h1>Acme Add-in</h1>
      <p className="info">
        Office web based add-in.
      </p>
      <div id="outlook-app">
        App connected to {host} on {platform}.
      </div>
      {elmMatches}
      <p>
        <button type="button" onClick={() => { throw new Error("Sentry Test Error") }}>Break the world</button>
      </p>
    </div>
  )
}

export default Sentry.withProfiler(Taskpane)
