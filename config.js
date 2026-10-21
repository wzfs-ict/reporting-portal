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
  clientId: "922cf5c9-1187-420a-af78-ec81f5f58227",   // "WZFS Reporting Portal"

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
    photos: "Student Photos",          // document library: <StudentID>.jpg, staff read only
    marksPrefix: "Marks ",
  },
  // SharePoint group that gets Contribute on the entry lists (Setup → Permissions)
  staffGroup: "Reporting Staff",

  // PDE: bobae goals live on a SEPARATE private site so bobaes never get access to the marks site.
  pde: {
    sitePath: "/sites/WZFSPDE",
    studentGroup: "PDE Students",
    advisorEntryGrades: ["G1", "G2", "G3", "G4", "G5"],   // advisors enter goals for these grades; the rest enter their own
  },

  // ── Term ─────────────────────────────────────────────────
  academicYear: "2026-27",
  semester: "Fall",               // Fall | Spring
  semesters: ["Fall", "Spring"],
  // Midterm = marks + remarks for the PTC sheets (not official). Final = the report card.
  // Each period keeps its own marks, so midterm entries are never overwritten by finals.
  period: "Midterm",
  periods: ["Midterm", "Final"],

  grades: ["G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8", "G9-1", "G9-2", "G10", "G11", "G12", "ZSA"],

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
  // Division heads: read-only Advisor view and PTC sheets for their own grades only.
  divisionHeads: {
    "xtobias@zhongshischool.org": ["G1", "G2", "G3", "G4", "G5"],   // Xochitl Tobias, Head of Primary
  },
  // Admin team: the only people who can generate report cards (Export tab). Admins are included automatically.
  reportTeam: [
    "rpark@zhongshischool.org",        // Ryan Park, Registrar
    "pgao@zhongshischool.org",         // Peter Gao, Administration
    "jbkim@zhongshischool.org",        // Keating Kim, Administrator
    "mkkim@zhongshischool.org",        // Andy Kim, VP Communications
  ],
  // Display names for the admin "View as" menu (advisors' names come from `advisors` below).
  names: {
    "glenj@zhongshischool.org": "Glen Joshua", "gjoshua@zhongshischool.org": "Glen Joshua (2nd account)",
    "rpark@zhongshischool.org": "Ryan Park", "pgao@zhongshischool.org": "Peter Gao", "jbkim@zhongshischool.org": "Keating Kim",
    "mkkim@zhongshischool.org": "Andy Kim", "xtobias@zhongshischool.org": "Xochitl Tobias", "asumeg-ang@zhongshischool.org": "Arsenio Sumeg-ang",
    "dlee@zhongshischool.org": "Daniel Lee", "mlee@zhongshischool.org": "Moody Lee", "rrussell@zhongshischool.org": "Roger Russell",
    "jpaa@zhongshischool.org": "Jonathan Pa-a", "agong@zhongshischool.org": "Alan Gong", "mparker@zhongshischool.org": "Mae Parker",
    "gparker@zhongshischool.org": "Gregg Parker", "randy@zhongshischool.org": "Randy Wang", "gliu@zhongshischool.org": "Grace Liu",
    "w_reacher@zhongshischool.org": "Reacher Williams",
    "bbirkenstock@zhongshischool.org": "Brendon Birkenstock", "eduarte@zhongshischool.org": "Eri Duarte", "kyang@zhongshischool.org": "Kimberly Yang",
  },
  // Teaching staff — pre-fills Setup → Permissions (the "Reporting Staff" group).
  staff: [
    // Primary School
    "xtobias@zhongshischool.org", "schoe@zhongshischool.org", "smuivah@zhongshischool.org",
    "msiao@zhongshischool.org", "mmabini@zhongshischool.org", "osubira@zhongshischool.org",
    "lrodriguez@zhongshischool.org", "bbirkenstock@zhongshischool.org",
    "eduarte@zhongshischool.org",      // Eri Duarte — PE (Athletics Director)
    "kyang@zhongshischool.org",        // Kimberly Yang — Choir / Orchestra
    // Upper School
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
    "G1":   { name: "Mrs. Xochitl Tobias",  co: "Ms. So Yeong Choe, Mrs. Grace Liu", email: "xtobias@zhongshischool.org, schoe@zhongshischool.org, gliu@zhongshischool.org" },
    "G2":   { name: "Ms. Singchon Muivah",  co: "Mrs. Alisa Wang",  email: "smuivah@zhongshischool.org, awang@zhongshischool.org" },
    "G3":   { name: "Mrs. Mailyn Siao",     co: "Mrs. Sasha Liu",   email: "msiao@zhongshischool.org, sliu@zhongshischool.org" },
    "G4":   { name: "Mrs. Margaret Mabini", co: "Mrs. Chloe Guo",   email: "mmabini@zhongshischool.org, cguo@zhongshischool.org" },
    "G5":   { name: "Mrs. Ola Fraser",      co: "Mr. Luis Rodriguez", email: "osubira@zhongshischool.org, lrodriguez@zhongshischool.org" },
    "G6":  { name: "Ms. Karen Bayangan",   co: "Mr. Stephen Feng", email: "kbayangan@zhongshischool.org, sfeng@zhongshischool.org" },
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
    "Certificate of Achievement",
    "Certificate of Distinction",
    "Certificate of Participation",
    "Teacher Award",
    "External Award",
    "Service Points",                // goes to the service points section of the report card
  ],
};
