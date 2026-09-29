// Demo data for the online (GitHub Pages) copy of dashboard.html.
// Names and numbers are made up. The real data, sent_log_data.js, is written by
// send_whatsapp.py on the computer that sends the messages and is never uploaded.
(function () {
  const d = new Date();
  const today = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  window.sentLogData = [
    { date: today, phone: "911234567890", occasion: "birthday", name: "FirstName SecondName",
      status: "sent", sent_at: `${today}T09:00:05`, message_id: "wamid.DEMO1" },
    { date: today, phone: "911234567891", occasion: "anniversary", name: "FirstName SecondName",
      status: "sent", sent_at: `${today}T09:00:07`, message_id: "wamid.DEMO2" },
    { date: today, phone: "911234567892", occasion: "birthday", name: "FirstName SecondName",
      status: "pending", sent_at: `${today}T09:00:09` },
  ];
})();
