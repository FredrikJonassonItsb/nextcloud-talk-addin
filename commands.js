// ====== konfiguration ======
const cfg = {
  baseUrl: localStorage.getItem("talkBaseUrl") || "https://demo.hubs.se",
  user:    localStorage.getItem("talkUser")    || "",
  token:   localStorage.getItem("talkToken")   || ""   // <‑ skapa app‑lösenord i Nextcloud
};

// ==== HTTP‑hjälpare ====
async function talkApi(path, method="GET", qs="") {
  const url = `${cfg.baseUrl}/ocs/v2.php/apps/spreed/api/v4/${path}${qs}`;
  const res = await fetch(url, {
    method,
    headers: {
      "OCS-APIRequest": "true",
      "Authorization" : "Basic " + btoa(`${cfg.user}:${cfg.token}`)
    }
  });
  const json = await res.json();
  if (json.ocs.meta.statuscode >= 400) throw new Error(json.ocs.meta.message);
  return json.ocs.data;
}

// ==== skapa nytt samtal & gör det publikt ====
async function generateTalkLink(roomName) {
  // 1. skapa rum
  const room = await talkApi(`room`, "POST", `?roomType=3&roomName=${encodeURIComponent(roomName)}`);
  const token = room.token;                                       // unikt id
  // 2. öppna för gäster
  await talkApi(`room/${token}/public`, "POST");
  return `${cfg.baseUrl}/call/${token}`;
}

// ==== ta bort Teams‑länk (förenklat) ====
function stripTeams(html) {
  return html.replace(/https:\/\/teams\.microsoft\.com\/[^"'\s]+/gi, "")
             .replace(/Join (Microsoft )?Teams Meeting[^<]*(<br>|$)/gi, "");
}

// ==== Knapparnas handlers ====
async function createTalkMeeting(event) {
  const subject = "Nextcloud Talk‑möte";
  const link    = await generateTalkLink(subject);
  Office.context.mailbox.displayNewAppointmentForm({
    subject,
    body: `Välkommen!<br><br><a href="${link}">Ans­lut till Nextcloud Talk‑möte</a>`
  });                                   // :contentReference[oaicite:2]{index=2}
  event.completed();
}

async function insertTalkLink(event) {
  const item = Office.context.mailbox.item;
  item.body.getAsync(Office.CoercionType.Html, async res => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) { event.completed(); return; }

    const cleaned = stripTeams(res.value);
    const link    = await generateTalkLink(item.subject || "Nextcloud Talk");
    const updated = cleaned + `<br><br><a href="${link}">Ans­lut till Nextcloud Talk‑möte</a><br>`;
    item.body.setAsync(updated, { coercionType: Office.CoercionType.Html }, () => event.completed());
  });
}

Office.actions.associate("createTalkMeeting", createTalkMeeting);
Office.actions.associate("insertTalkLink",   insertTalkLink);
