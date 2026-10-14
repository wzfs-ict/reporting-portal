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
  admins: ["glenj@zhongshischool.org"],
  // Read-only across all grades (they must also be site Owners in SharePoint)
  leadership: [
    "asumeg-ang@zhongshischool.org",
    "dlee@zhongshischool.org",
    "rrussell@zhongshischool.org",
    "mkkim@zhongshischool.org",
  ],
  // Grade → class advisor (from AAO "26 Fall Bobae's Picture List", 15 Sept 2026).
  // `name` prints on the report card. `email` may list the advisor AND co-advisor,
  // comma-separated — everyone listed can see and correct the whole grade.
  // Blank email = not yet confirmed. Edit on the Setup tab.
  advisors: {
    "G6":   { email: "", name: "Ms. Karen Bayangan", co: "Mr. Stephen Feng" },
    "G7":   { email: "", name: "Mrs. Prapti Patole", co: "Mr. Richard Wang" },
    "G8":   { email: "b_reacher@zhongshischool.org", name: "Mrs. Bijily Reacher", co: "Mrs. Judy Wang" },
    "G9-1": { email: "glenj@zhongshischool.org", name: "Mrs. Keren Joshua", co: "Mr. Glen Joshua" },
    "G9-2": { email: "", name: "Mr. Arnold Bagapuro", co: "Mr. Austin Yu" },
    "G10":  { email: "", name: "Mr. Michael Bardinas", co: "Mrs. Winnie Li" },
    "G11":  { email: "jpatole@zhongshischool.org", name: "Mr. Jayesh Patole", co: "Mrs. Summer Xia" },
    "G12":  { email: "babuan@zhongshischool.org", name: "Ms. Beverly Abuan", co: "Mrs. Ruth Xia" },
    "ZSA":  { email: "", name: "Mrs. May Cho", co: "" },
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
