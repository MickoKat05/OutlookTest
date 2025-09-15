Office.onReady(() => {
  const item = Office.context.mailbox.item;

  Office.context.mailbox.item.to.getAsync((result) => {
    const recipients = result.value.map(r => r.emailAddress.toLowerCase());
    if (recipients.includes("support@firstmajestic.com")) {
      const token = "hleu0jKW7DdsnfumG2IO";
      const headers = {
        "Authorization": "Basic " + btoa(token + ":"),
        "Content-Type": "application/json"
      };

      Promise.all([
        fetch("https://firstmajestic.freshservice.com/api/v2/departments", { headers }).then(res => res.json()),
        fetch("https://firstmajestic.freshservice.com/api/v2/agents", { headers }).then(res => res.json())
      ]).then(([departmentsData, agentsData]) => {
        const departments = departmentsData.departments || [];
        const agents = agentsData.agents || [];

        const departmentOptions = departments.map(d => `<option>${d.name}</option>`).join("");
        const agentOptions = agents.map(a => `<option>${a.contact.name}</option>`).join("");

        const htmlForm = `
          <div style="border:1px solid #ccc; padding:10px;">
            <h3>Formulario de Ticket</h3>
            <label>Agente:</label>
            <select>${agentOptions}</select><br><br>
            <label>Departamento:</label>
            <select>${departmentOptions}</select><br><br>
            <label>Descripción:</label><br>
            <textarea rows="4" cols="50"></textarea>
          </div>
        `;

        item.body.setAsync(htmlForm, { coercionType: Office.CoercionType.Html });
      });
    }
  });
});
