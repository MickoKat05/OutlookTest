Office.initialize = function () {};

function insertFreshserviceFields(event) {
  Office.context.mailbox.item.to.getAsync(function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const recipients = result.value.map(r => r.emailAddress.toLowerCase());
      if (recipients.includes("support@firstmajestic.com")) {
        fetchDepartmentsAndAgents().then(({ departments, agents }) => {
          const html = `
            <p><strong>Descripción:</strong><br/><textarea rows="4" cols="50"></textarea></p>
            <p><strong>Departamento:</strong><br/>
              <select>${departments.map(d => `<option>${d}</option>`).join("")}</select>
            </p>
            <p><strong>Agente:</strong><br/>
              <select>${agents.map(a => `<option>${a}</option>`).join("")}</select>
            </p>
          `;
          Office.context.mailbox.item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, function (asyncResult) {
            event.completed();
          });
        });
      } else {
        event.completed();
      }
    } else {
      event.completed();
    }
  });
}

function fetchDepartmentsAndAgents() {
  const token = "hleu0jKW7DdsnfumG2IO";
  const headers = { "Authorization": "Basic " + btoa(token + ":X") };
  const deptUrl = "https://firstmajestic.freshservice.com/api/v2/departments";
  const agentUrl = "https://firstmajestic.freshservice.com/api/v2/agents";

  return Promise.all([
    fetch(deptUrl, { headers }).then(res => res.json()).then(data => data.departments.map(d => d.name)),
    fetch(agentUrl, { headers }).then(res => res.json()).then(data => data.agents.map(a => a.name))
  ]).then(([departments, agents]) => ({ departments, agents }));
}
