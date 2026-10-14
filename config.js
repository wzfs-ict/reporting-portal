// ─────────────────────────────────────────────────────────────
//  WZFS Reporting Portal — settings
//  This is the ONLY file you normally need to edit.
//  Email addresses must be lower-case school accounts.
//  Once someone saves on the Setup tab, the people/term saved there
//  replace the starting values below.
// ─────────────────────────────────────────────────────────────
window.APP_CONFIG = {
  appName: "WZFS Reporting",
  schoolName: "Weihai Zhongshi Foreign School",
  version: "V1",

  // Microsoft Entra app registration (README step 2)
  tenantId: "59aaa89c-ebfe-417a-bcdf-21065b509cf4",
  clientId: "PASTE-APPLICATION-CLIENT-ID-HERE",

  // The private SharePoint site that stores the data (README step 1)
  siteHostname: "weihaizhongshi.sharepoint.com",
  sitePath: "/sites/WZFSReporting",

  // List names. Marks lists are one per grade: "Marks G6", "Marks G7", …
  lists: {
    students: "Students",
    subjects: "Subjects",
    catalogue: "Award Catalogue",
    awards: "Awards",
    comments: "Advisor Comments",
    settings: "Reporting Settings",
    marksPrefix: "Marks ",
  },
  // SharePoint group that gets Contribute on the entry lists (Setup → Permissions)
  staffGroup: "Reporting Staff",

  // ── Term ─────────────────────────────────────────────────
  academicYear: "2026-27",
  semester: "Fall",               // Fall | Spring
  semesters: ["Fall", "Spring"],

  grades: ["G6", "G7", "G8", "G9-1", "G9-2", "G10", "G11", "G12", "ZSA"],

  // ── People ───────────────────────────────────────────────
  // Accounts verified against AllUsersM365.csv (Oct 2026). Some differ from the
  // staff list: Bijily = b_reacher@, Andy Kim = mkkim@, Reacher Williams = w_reacher@.
  admins: ["glenj@zhongshischool.org", "gjoshua@zhongshischool.org"],   // Glen has both accounts
  // Read-only across all grades (they must also be site Owners in SharePoint)
  leadership: [
    "mlee@zhongshischool.org",         // Moody Lee, Chairman
    "asumeg-ang@zhongshischool.org",   // Arsenio Sumeg-ang, Principal
    "dlee@zhongshischool.org",         // Daniel Lee, Vice President
    "mkkim@zhongshischool.org",        // Andy Kim, VP Communications
    "rrussell@zhongshischool.org",     // Roger Russell
  ],
  // Upper School teaching staff — pre-fills Setup → Permissions (the "Reporting Staff" group).
  staff: [
    "abagapuro@zhongshischool.org", "b_reacher@zhongshischool.org", "jpatole@zhongshischool.org",
    "jpaa@zhongshischool.org", "ayu@zhongshischool.org", "agong@zhongshischool.org",
    "wli@zhongshischool.org", "gliu@zhongshischool.org", "mparker@zhongshischool.org",
    "gparker@zhongshischool.org", "kbayangan@zhongshischool.org", "sxia@zhongshischool.org",
    "randy@zhongshischool.org", "ppatole@zhongshischool.org", "rxia@zhongshischool.org",
    "sliu@zhongshischool.org", "jwang@zhongshischool.org", "kjoshua@zhongshischool.org",
    "babuan@zhongshischool.org", "awang@zhongshischool.org", "mbardinas@zhongshischool.org",
    "sfeng@zhongshischool.org", "rqi@zhongshischool.org", "cguo@zhongshischool.org",
    "w_reacher@zhongshischool.org", "mcho@zhongshischool.org",
    "glenj@zhongshischool.org", "gjoshua@zhongshischool.org",
  ],
  // Grade → class advisor (from AAO "26 Fall Bobae's Picture List", 15 Sept 2026).
  // `name` prints on the report card. `email` may list the advisor AND co-advisor,
  // comma-separated — everyone listed can see and correct the whole grade.
  // Blank email = not yet confirmed. Edit on the Setup tab.
  advisors: {
    "G6":   { name: "Ms. Karen Bayangan",   co: "Mr. Stephen Feng", email: "kbayangan@zhongshischool.org, sfeng@zhongshischool.org" },
    "G7":   { name: "Mrs. Prapti Patole",   co: "Mr. Richard Wang", email: "ppatole@zhongshischool.org, rqi@zhongshischool.org" },
    "G8":   { name: "Mrs. Bijily Reacher",  co: "Mrs. Judy Wang",   email: "b_reacher@zhongshischool.org, jwang@zhongshischool.org" },
    "G9-1": { name: "Mrs. Keren Joshua",    co: "Mr. Glen Joshua",  email: "kjoshua@zhongshischool.org, glenj@zhongshischool.org, gjoshua@zhongshischool.org" },
    "G9-2": { name: "Mr. Arnold Bagapuro",  co: "Mr. Austin Yu",    email: "abagapuro@zhongshischool.org, ayu@zhongshischool.org" },
    "G10":  { name: "Mr. Michael Bardinas", co: "Mrs. Winnie Li",   email: "mbardinas@zhongshischool.org, wli@zhongshischool.org" },
    "G11":  { name: "Mr. Jayesh Patole",    co: "Mrs. Summer Xia",  email: "jpatole@zhongshischool.org, sxia@zhongshischool.org" },
    "G12":  { name: "Ms. Beverly Abuan",    co: "Mrs. Ruth Xia",    email: "babuan@zhongshischool.org, rxia@zhongshischool.org" },
    "ZSA":  { name: "Mrs. May Cho",         co: "",                 email: "mcho@zhongshischool.org" },
  },

  // ── Scales ───────────────────────────────────────────────
  behaviourScale: [
    { value: 1, label: "Excellent" },
    { value: 2, label: "Good" },
    { value: 3, label: "Satisfactory" },
    { value: 4, label: "Needs Improvement" },
  ],
  remarkMaxLength: 300,
  certificateTypes: [
    "Certificate of Recognition",
    "Certificate of Distinction",
    "Certificate of Participation",
    "Teacher Award",
  ],
};
