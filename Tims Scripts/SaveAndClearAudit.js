//[group=Tims Scripts]
!INC Local Scripts.EAConstants-JScript

/*
 * Script Name: 
 * Author: 
 * Purpose: 
 * Date: 
 */
 
function getMonth(date) {
 month = date.getMonth() + 1;
 if (month < 10) {
  return "0" + month;
 } else {
  return month;
 }
}
 
function main()
{
 today = new Date();
 date = "" + today.getFullYear() + getMonth(today) + today.getDate();
 filename = "H:\\_Projecten\\EA\\Sparx\\EAAuditLog-" + date + ".xml";
 if (Repository.SaveAuditLogs(filename, null, null)) {
  Session.Output("auditlogs saved to " + filename);
  if (Repository.ClearAuditLogs(null, null)) {
   Session.Output("auditlogs cleared");
  } else {
   Session.Output("problem clearing auditlogs");
  }
 } else {
  Session.Output("problem saving auditlogs, not cleared");
 }
}

main();