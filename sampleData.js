/**
 * Default Seed & Sample Data
 * Extracted directly from company Excel files: PH497_2026-08-05.xlsx & PH497_2026-09-05.xlsx
 */

window.DATA_VERSION = "2026-09-05-v3";

window.DEFAULT_PROJECTS = {
  "PH497-SG1": {
    projectNumber: "PH497",
    projectName: "SG1",
    displayName: "PH497-SG1",
    client: "Design Department",
    projectLead: "Pranav",
    projectInCharge: "Rejoe",
    bimCoordinator: "Ajmal",
    projectDc: "Romeo",
    procurement: "Osama",
    estimation: "Abdul",
    plannedStartDate: "01/01/2026",
    plannedEndDate: "06/06/2026",
    actualStartDate: "04/04/2026",
    presentDate: "05/09/2026",
    periodOfExtension: "3 Months",
    modelers: [
      { name: "OSR 1 (Faisal Khan)", role: "Senior BIM Modeler", dateOfJoin: "15/01/2026", active: true },
      { name: "OSR 2 (Naveen Kumar)", role: "Revit Architectural Modeler", dateOfJoin: "01/02/2026", active: true },
      { name: "OSR 3 (Rajesh Pillai)", role: "Structural Modeler", dateOfJoin: "15/02/2026", active: true },
      { name: "OSR 4 (Daniel Thomas)", role: "MEP Modeler", dateOfJoin: "01/03/2026", active: true },
      { name: "OSR 5 (Anand Varma)", role: "Junior Modeler", dateOfJoin: "15/03/2026", active: false }
    ],
    notes: [
      "Large number of drawings are pending due to a high number of open technical queries with the client.",
      "Ground floor & Zone 01 drawings were resubmitted with client comments on 01/09/2026.",
      "Structural REV-B modifications completed for ML-000 foundation layouts."
    ],
    history: {
      "2026-08-05": {
        fileName: "PH497_2026-08-05.xlsx",
        dateOfReport: "05/08/2026",
        presentDate: "05/08/2026",
        plannedEndDate: "06/06/2026",
        periodOfExtension: "2 Months",
        monthYearLabel: "Aug, 26",
        totalDwg: 194,
        resourceCount: 5,
        statusSummary: {
          approved: 160,
          approvedWithResub: 0,
          rejected: 5,
          underPrep: 1,
          pending: 17,
          submitted: 11,
          notStarted: 0
        },
        completionPct: 82.5,
        drawings: [{"row":2,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915001"},{"row":3,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915002"},{"row":4,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915003"},{"row":5,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-000-AR-915001"},{"row":6,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-000-AR-915002"},{"row":7,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915004"},{"row":8,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915005"},{"row":9,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915006"},{"row":10,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915007"},{"row":11,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915008"},{"row":12,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915009"},{"row":13,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915010"},{"row":14,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915011"},{"row":15,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-000-AR-915003"},{"row":16,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z01-AR-215001"},{"row":17,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z02-AR-215001"},{"row":18,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z03-AR-215001"},{"row":19,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z04-AR-215001"},{"row":20,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-SWID-L1-Z05-AR-215001"},{"row":21,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z06-AR-215001"},{"row":22,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z07-AR-215001"},{"row":23,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z08-AR-215001"},{"row":24,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z09-AR-215001"},{"row":25,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z10-AR-215001"},{"row":26,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z01-AR-215001"},{"row":27,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z02-AR-215001"},{"row":28,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z03-AR-215001"},{"row":29,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z04-AR-215001"},{"row":30,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z05-AR-215001"},{"row":31,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z07-AR-215001"},{"row":32,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-AN01-GF-EX1-AR-515001"},{"row":33,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-AN02-GF-EX1-AR-515001"},{"row":34,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-AN03-GF-EX4-AR-515001"},{"row":35,"status":"Approved","drawingNo":"SG1-AC1-SD-AN04-GF-EX6-AR-515001"},{"row":36,"status":"Approved","drawingNo":"SG1-AC1-SD-GH01-GF-EX1-AR-515001"},{"row":37,"status":"Approved","drawingNo":"SG1-AC1-SD-GH02-GF-EX7-AR-515001"},{"row":38,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-MV01-GF-EX3-AR-515001"},{"row":39,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-MV02-GF-000-AR-515001"},{"row":40,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-MV03-GF-000-AR-515001"},{"row":41,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-MV04-GF-EX9-AR-515001"},{"row":42,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z01-AR-215001"},{"row":43,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z02-AR-215001"},{"row":44,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z03-AR-215001"},{"row":45,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z04-AR-215001"},{"row":46,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z04-AR-215002"},{"row":47,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z05-AR-215001"},{"row":48,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z06-AR-215001"},{"row":49,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z06-AR-215002"},{"row":50,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z07-AR-215001"},{"row":51,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z08-AR-215001"},{"row":52,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z01-AR-215002"},{"row":53,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z01-AR-215003"},{"row":54,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z01-AR-215004"},{"row":55,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z01-AR-215005"},{"row":56,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z02-AR-215002"},{"row":57,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z02-AR-215003"},{"row":58,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z02-AR-215004"},{"row":59,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z02-AR-215005"},{"row":60,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z03-AR-215002"},{"row":61,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z03-AR-215003"},{"row":62,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z03-AR-215004"},{"row":63,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z03-AR-215005"},{"row":64,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z04-AR-215002"},{"row":65,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z04-AR-215003"},{"row":66,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z04-AR-215004"},{"row":67,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z04-AR-215005"},{"row":68,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z05-AR-215002"},{"row":69,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z05-AR-215003"},{"row":70,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z05-AR-215004"},{"row":71,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z05-AR-215005"},{"row":72,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z07-AR-215002"},{"row":73,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z07-AR-215003"},{"row":74,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z07-AR-215004"},{"row":75,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z07-AR-215005"},{"row":76,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-OFF-AR-215001"},{"row":77,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-OFF-AR-515002"},{"row":78,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-OFF-AR-515001"},{"row":79,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515009"},{"row":80,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515001"},{"row":81,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z01-AR-215003"},{"row":82,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z02-AR-215003"},{"row":83,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z03-AR-215003"},{"row":84,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z04-AR-215003"},{"row":85,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z05-AR-215003"},{"row":86,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z06-AR-215003"},{"row":87,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z07-AR-215003"},{"row":88,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z08-AR-215003"},{"row":89,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z09-AR-215003"},{"row":90,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z10-AR-215003"},{"row":91,"status":"Approved","drawingNo":"SG1-AC1-SD-PR01-GF-EX1-AR-215001"},{"row":92,"status":"Approved","drawingNo":"SG1-AC1-SD-PR01-GF-EX1-AR-215002"},{"row":93,"status":"Approved","drawingNo":"SG1-AC1-SD-PR01-GF-EX1-AR-515003"},{"row":94,"status":"Approved","drawingNo":"SG1-AC1-SD-PR02-GF-EX4-AR-215001"},{"row":95,"status":"Approved","drawingNo":"SG1-AC1-SD-PR02-GF-EX4-AR-215002"},{"row":96,"status":"Approved","drawingNo":"SG1-AC1-SD-PR02-GF-EX4-AR-515003"},{"row":97,"status":"Approved","drawingNo":"SG1-AC1-SD-PR03-GF-EX7-AR-215001"},{"row":98,"status":"Approved","drawingNo":"SG1-AC1-SD-PR03-GF-EX7-AR-215002"},{"row":99,"status":"Approved","drawingNo":"SG1-AC1-SD-PR03-GF-EX7-AR-215003"},{"row":100,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z03-AR-515003"},{"row":101,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z03-AR-515004"},{"row":102,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z01-AR-215004"},{"row":103,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z01-AR-215005"},{"row":104,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z03-AR-215001"},{"row":105,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z04-AR-515001"},{"row":106,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z04-AR-515002"},{"row":107,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z07-AR-215001"},{"row":108,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z08-AR-215001"},{"row":109,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z08-AR-215002"},{"row":110,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z01-AR-215001"},{"row":111,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z01-AR-215002"},{"row":112,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z02-AR-515001"},{"row":113,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z05-AR-215001"},{"row":114,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z05-AR-215002"},{"row":115,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z06-AR-215001"},{"row":116,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z06-AR-215002"},{"row":117,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-SG1A-ML-Z10-AR-515001"},{"row":118,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z09-AR-515001"},{"row":119,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z09-AR-515002"},{"row":120,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z01-AR-215001"},{"row":121,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z02-AR-215001"},{"row":122,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z03-AR-215001"},{"row":123,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z04-AR-215001"},{"row":124,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z05-AR-215001"},{"row":125,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z06-AR-215001"},{"row":126,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z07-AR-215001"},{"row":127,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z08-AR-215001"},{"row":128,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z09-AR-215001"},{"row":129,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z10-AR-215001"},{"row":130,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z01-AR-215002"},{"row":131,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z02-AR-215002"},{"row":132,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z03-AR-215002"},{"row":133,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z04-AR-215002"},{"row":134,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z05-AR-215002"},{"row":135,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z06-AR-515002"},{"row":136,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z07-AR-215002"},{"row":137,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z08-AR-215002"},{"row":138,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z09-AR-215002"},{"row":139,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z10-AR-215002"},{"row":140,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-RF-Z03-AR-515001"},{"row":141,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-RF-Z03-AR-515002"},{"row":142,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-RF-Z01-AR-515003"},{"row":143,"status":"Rejected","drawingNo":"SG1-AC1-SD-SG1A-RF-Z04-AR-215001"},{"row":144,"status":"Rejected","drawingNo":"SG1-AC1-SD-SG1A-RF-Z05-AR-215001"},{"row":145,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-RF-Z08-AR-515001"},{"row":146,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-SG1A-RF-Z09-AR-515001"},{"row":147,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-RF-Z01-AR-515002"},{"row":148,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-SG1A-RF-Z09-AR-215002"},{"row":149,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-SG1A-RF-Z09-AR-215003"},{"row":150,"status":"Rejected","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915001"},{"row":151,"status":"Rejected","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915002"},{"row":152,"status":"Under Preparation","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915003"},{"row":153,"status":"Rejected","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915004"},{"row":154,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915005"},{"row":155,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915006"},{"row":156,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915007"},{"row":157,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915008"},{"row":158,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515001"},{"row":159,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515002"},{"row":160,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515003"},{"row":161,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515004"},{"row":162,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515005"},{"row":163,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515006"},{"row":164,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515007"},{"row":165,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX3-AR-215001"},{"row":166,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX4-AR-515001"},{"row":167,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX4-AR-515002"},{"row":168,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX4-AR-515003"},{"row":169,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX4-AR-515004"},{"row":170,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-EX4-AR-515005"},{"row":171,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515002"},{"row":172,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515003"},{"row":173,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515004"},{"row":174,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515005"},{"row":175,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515006"},{"row":176,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515007"},{"row":177,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515008"},{"row":178,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515009"},{"row":179,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515010"},{"row":180,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515001"},{"row":181,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515002"},{"row":182,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515003"},{"row":183,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515004"},{"row":184,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515005"},{"row":185,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515006"},{"row":186,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515007"},{"row":187,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515008"},{"row":188,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515009"},{"row":189,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515010"},{"row":190,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515011"},{"row":191,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515012"},{"row":192,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515013"},{"row":193,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515014"},{"row":194,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515015"},{"row":195,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515016"}]
      },
      "2026-09-05": {
        fileName: "PH497_2026-09-05.xlsx",
        dateOfReport: "05/09/2026",
        presentDate: "05/09/2026",
        plannedEndDate: "06/06/2026",
        periodOfExtension: "3 Months",
        monthYearLabel: "Sep, 26",
        totalDwg: 194,
        resourceCount: 4,
        statusSummary: {
          approved: 134,
          approvedWithResub: 0,
          rejected: 5,
          underPrep: 1,
          pending: 43,
          submitted: 11,
          notStarted: 0
        },
        completionPct: 69.1,
        drawings: [{"row":2,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915001"},{"row":3,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915002"},{"row":4,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915003"},{"row":5,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-000-AR-915001"},{"row":6,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-000-AR-915002"},{"row":7,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915004"},{"row":8,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915005"},{"row":9,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915006"},{"row":10,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915007"},{"row":11,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915008"},{"row":12,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915009"},{"row":13,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915010"},{"row":14,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-000-AR-915011"},{"row":15,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-000-AR-915003"},{"row":16,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z01-AR-215001"},{"row":17,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z02-AR-215001"},{"row":18,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z03-AR-215001"},{"row":19,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z04-AR-215001"},{"row":20,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-SWID-L1-Z05-AR-215001"},{"row":21,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z06-AR-215001"},{"row":22,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z07-AR-215001"},{"row":23,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z08-AR-215001"},{"row":24,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z09-AR-215001"},{"row":25,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z10-AR-215001"},{"row":26,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z01-AR-215001"},{"row":27,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z02-AR-215001"},{"row":28,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z03-AR-215001"},{"row":29,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z04-AR-215001"},{"row":30,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z05-AR-215001"},{"row":31,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z07-AR-215001"},{"row":32,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-AN01-GF-EX1-AR-515001"},{"row":33,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-AN02-GF-EX1-AR-515001"},{"row":34,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-AN03-GF-EX4-AR-515001"},{"row":35,"status":"Approved","drawingNo":"SG1-AC1-SD-AN04-GF-EX6-AR-515001"},{"row":36,"status":"Approved","drawingNo":"SG1-AC1-SD-GH01-GF-EX1-AR-515001"},{"row":37,"status":"Approved","drawingNo":"SG1-AC1-SD-GH02-GF-EX7-AR-515001"},{"row":38,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-MV01-GF-EX3-AR-515001"},{"row":39,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-MV02-GF-000-AR-515001"},{"row":40,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-MV03-GF-000-AR-515001"},{"row":41,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-MV04-GF-EX9-AR-515001"},{"row":42,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z01-AR-215001"},{"row":43,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z02-AR-215001"},{"row":44,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z03-AR-215001"},{"row":45,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z04-AR-215001"},{"row":46,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z04-AR-215002"},{"row":47,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z05-AR-215001"},{"row":48,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z06-AR-215001"},{"row":49,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z06-AR-215002"},{"row":50,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z07-AR-215001"},{"row":51,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-Z08-AR-215001"},{"row":52,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z01-AR-215002"},{"row":53,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z01-AR-215003"},{"row":54,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z01-AR-215004"},{"row":55,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z01-AR-215005"},{"row":56,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z02-AR-215002"},{"row":57,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z02-AR-215003"},{"row":58,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z02-AR-215004"},{"row":59,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z02-AR-215005"},{"row":60,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z03-AR-215002"},{"row":61,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z03-AR-215003"},{"row":62,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z03-AR-215004"},{"row":63,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z03-AR-215005"},{"row":64,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z04-AR-215002"},{"row":65,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z04-AR-215003"},{"row":66,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z04-AR-215004"},{"row":67,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z04-AR-215005"},{"row":68,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z05-AR-215002"},{"row":69,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z05-AR-215003"},{"row":70,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z05-AR-215004"},{"row":71,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z05-AR-215005"},{"row":72,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z07-AR-215002"},{"row":73,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z07-AR-215003"},{"row":74,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z07-AR-215004"},{"row":75,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-L1-Z07-AR-215005"},{"row":76,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-OFF-AR-215001"},{"row":77,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-GF-OFF-AR-515002"},{"row":78,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-ML-OFF-AR-515001"},{"row":79,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515009"},{"row":80,"status":"Approved","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515001"},{"row":81,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z01-AR-215003"},{"row":82,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z02-AR-215003"},{"row":83,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z03-AR-215003"},{"row":84,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z04-AR-215003"},{"row":85,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z05-AR-215003"},{"row":86,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z06-AR-215003"},{"row":87,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z07-AR-215003"},{"row":88,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z08-AR-215003"},{"row":89,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z09-AR-215003"},{"row":90,"status":"Approved","drawingNo":"SG1-AC1-SD-SWID-RF-Z10-AR-215003"},{"row":91,"status":"Approved","drawingNo":"SG1-AC1-SD-PR01-GF-EX1-AR-215001"},{"row":92,"status":"Approved","drawingNo":"SG1-AC1-SD-PR01-GF-EX1-AR-215002"},{"row":93,"status":"Approved","drawingNo":"SG1-AC1-SD-PR01-GF-EX1-AR-515003"},{"row":94,"status":"Approved","drawingNo":"SG1-AC1-SD-PR02-GF-EX4-AR-215001"},{"row":95,"status":"Approved","drawingNo":"SG1-AC1-SD-PR02-GF-EX4-AR-215002"},{"row":96,"status":"Approved","drawingNo":"SG1-AC1-SD-PR02-GF-EX4-AR-515003"},{"row":97,"status":"Approved","drawingNo":"SG1-AC1-SD-PR03-GF-EX7-AR-215001"},{"row":98,"status":"Approved","drawingNo":"SG1-AC1-SD-PR03-GF-EX7-AR-215002"},{"row":99,"status":"Approved","drawingNo":"SG1-AC1-SD-PR03-GF-EX7-AR-215003"},{"row":100,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z03-AR-515003"},{"row":101,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z03-AR-515004"},{"row":102,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z01-AR-215004"},{"row":103,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z01-AR-215005"},{"row":104,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z03-AR-215001"},{"row":105,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z04-AR-515001"},{"row":106,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z04-AR-515002"},{"row":107,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z07-AR-215001"},{"row":108,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z08-AR-215001"},{"row":109,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z08-AR-215002"},{"row":110,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z01-AR-215001"},{"row":111,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z01-AR-215002"},{"row":112,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z02-AR-515001"},{"row":113,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z05-AR-215001"},{"row":114,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z05-AR-215002"},{"row":115,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z06-AR-215001"},{"row":116,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z06-AR-215002"},{"row":117,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-SG1A-ML-Z10-AR-515001"},{"row":118,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z09-AR-515001"},{"row":119,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-ML-Z09-AR-515002"},{"row":120,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z01-AR-215001"},{"row":121,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z02-AR-215001"},{"row":122,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z03-AR-215001"},{"row":123,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z04-AR-215001"},{"row":124,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z05-AR-215001"},{"row":125,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z06-AR-215001"},{"row":126,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z07-AR-215001"},{"row":127,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z08-AR-215001"},{"row":128,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z09-AR-215001"},{"row":129,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z10-AR-215001"},{"row":130,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z01-AR-215002"},{"row":131,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z02-AR-215002"},{"row":132,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z03-AR-215002"},{"row":133,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z04-AR-215002"},{"row":134,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z05-AR-215002"},{"row":135,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z06-AR-515002"},{"row":136,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z07-AR-215002"},{"row":137,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z08-AR-215002"},{"row":138,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z09-AR-215002"},{"row":139,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-GF-Z10-AR-215002"},{"row":140,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-RF-Z03-AR-515001"},{"row":141,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-RF-Z03-AR-515002"},{"row":142,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-RF-Z01-AR-515003"},{"row":143,"status":"Rejected","drawingNo":"SG1-AC1-SD-SG1A-RF-Z04-AR-215001"},{"row":144,"status":"Rejected","drawingNo":"SG1-AC1-SD-SG1A-RF-Z05-AR-215001"},{"row":145,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-RF-Z08-AR-515001"},{"row":146,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-SG1A-RF-Z09-AR-515001"},{"row":147,"status":"Approved","drawingNo":"SG1-AC1-SD-SG1A-RF-Z01-AR-515002"},{"row":148,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-SG1A-RF-Z09-AR-215002"},{"row":149,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-SG1A-RF-Z09-AR-215003"},{"row":150,"status":"Rejected","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915001"},{"row":151,"status":"Rejected","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915002"},{"row":152,"status":"Under Preparation","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915003"},{"row":153,"status":"Rejected","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915004"},{"row":154,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915005"},{"row":155,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915006"},{"row":156,"status":"Submitted/Under Review","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915007"},{"row":157,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-915008"},{"row":158,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515001"},{"row":159,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515002"},{"row":160,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515003"},{"row":161,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515004"},{"row":162,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515005"},{"row":163,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515006"},{"row":164,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX1-AR-515007"},{"row":165,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX3-AR-215001"},{"row":166,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX4-AR-515001"},{"row":167,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX4-AR-515002"},{"row":168,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX4-AR-515003"},{"row":169,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX4-AR-515004"},{"row":170,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX4-AR-515005"},{"row":171,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515002"},{"row":172,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515003"},{"row":173,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515004"},{"row":174,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515005"},{"row":175,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515006"},{"row":176,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515007"},{"row":177,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515008"},{"row":178,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515009"},{"row":179,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-EX7-AR-515010"},{"row":180,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515001"},{"row":181,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515002"},{"row":182,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515003"},{"row":183,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515004"},{"row":184,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515005"},{"row":185,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515006"},{"row":186,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515007"},{"row":187,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515008"},{"row":188,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515009"},{"row":189,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515010"},{"row":190,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515011"},{"row":191,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515012"},{"row":192,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515013"},{"row":193,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515014"},{"row":194,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515015"},{"row":195,"status":"Pending(Notes)","drawingNo":"SG1-AC1-SD-EXWS-GF-000-AR-515016"}]
      }
    }
  },
  "PH498-SG2": {
    projectNumber: "PH498",
    projectName: "SG2 Commercial",
    displayName: "PH498-SG2",
    projectLead: "Pranav",
    projectInCharge: "Rejoe",
    bimCoordinator: "Ajmal",
    projectDc: "Romeo",
    procurement: "Osama",
    estimation: "Abdul",
    plannedStartDate: "15/02/2026",
    plannedEndDate: "15/08/2026",
    actualStartDate: "01/03/2026",
    presentDate: "05/09/2026",
    periodOfExtension: "1 Month",
    modelers: [
      { name: "OSR 1 (Vikram Patel)", role: "Senior BIM Modeler", dateOfJoin: "15/02/2026", active: true },
      { name: "OSR 2 (Arjun Das)", role: "Architectural Modeler", dateOfJoin: "01/03/2026", active: true }
    ],
    notes: [
      "Client requested revised mezzanine layouts; pending architect sign-off.",
      "Site survey differences identified on Grid F; technical query #14 submitted."
    ],
    history: {
      "2026-08-05": {
        fileName: "PH498_2026-08-05.xlsx",
        dateOfReport: "05/08/2026",
        presentDate: "05/08/2026",
        plannedEndDate: "15/08/2026",
        periodOfExtension: "0 Months",
        monthYearLabel: "Aug, 26",
        totalDwg: 140,
        resourceCount: 3,
        statusSummary: {
          approved: 70,
          approvedWithResub: 10,
          rejected: 4,
          underPrep: 12,
          pending: 15,
          submitted: 10,
          notStarted: 19
        },
        completionPct: 50.0,
        drawings: []
      },
      "2026-09-05": {
        fileName: "PH498_2026-09-05.xlsx",
        dateOfReport: "05/09/2026",
        presentDate: "05/09/2026",
        plannedEndDate: "15/08/2026",
        periodOfExtension: "1 Month",
        monthYearLabel: "Sep, 26",
        totalDwg: 140,
        resourceCount: 2,
        statusSummary: {
          approved: 95,
          approvedWithResub: 12,
          rejected: 3,
          underPrep: 15,
          pending: 10,
          submitted: 0,
          notStarted: 5
        },
        completionPct: 67.9,
        drawings: []
      }
    }
  },
  "PH502-DXB": {
    projectNumber: "PH502",
    projectName: "DXB Tower A",
    displayName: "PH502-DXB",
    projectLead: "Rejoe",
    projectInCharge: "Rejoe",
    bimCoordinator: "Ajmal",
    projectDc: "Romeo",
    procurement: "Osama",
    estimation: "Abdul",
    plannedStartDate: "01/04/2026",
    plannedEndDate: "30/09/2026",
    actualStartDate: "10/04/2026",
    presentDate: "05/09/2026",
    periodOfExtension: "0 Months",
    modelers: [
      { name: "OSR 1 (Suresh Nair)", role: "BIM Coordinator", dateOfJoin: "10/04/2026", active: true },
      { name: "OSR 2 (Farooq Al-Mansoor)", role: "MEP Modeler", dateOfJoin: "20/04/2026", active: true }
    ],
    notes: [
      "All podium drawings approved with code A.",
      "Tower facade package under preparation for submittal next week."
    ],
    history: {
      "2026-08-05": {
        fileName: "PH502_2026-08-05.xlsx",
        dateOfReport: "05/08/2026",
        presentDate: "05/08/2026",
        plannedEndDate: "30/09/2026",
        periodOfExtension: "0 Months",
        monthYearLabel: "Aug, 26",
        totalDwg: 88,
        resourceCount: 3,
        statusSummary: {
          approved: 45,
          approvedWithResub: 5,
          rejected: 2,
          underPrep: 15,
          pending: 12,
          submitted: 0,
          notStarted: 9
        },
        completionPct: 51.1,
        drawings: []
      },
      "2026-09-05": {
        fileName: "PH502_2026-09-05.xlsx",
        dateOfReport: "05/09/2026",
        presentDate: "05/09/2026",
        plannedEndDate: "30/09/2026",
        periodOfExtension: "0 Months",
        monthYearLabel: "Sep, 26",
        totalDwg: 88,
        resourceCount: 2,
        statusSummary: {
          approved: 60,
          approvedWithResub: 5,
          rejected: 2,
          underPrep: 8,
          pending: 8,
          submitted: 0,
          notStarted: 5
        },
        completionPct: 68.2,
        drawings: []
      }
    }
  }
};