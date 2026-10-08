export type RegisterField = {
  key: string;
  label: string;
  type: "text" | "number" | "date" | "select" | "textarea";
  group: string;
  required?: boolean;
  placeholder?: string;
  options?: string[];
  step?: string;
  unit?: string;
};

export type RegisterModule = {
  id: string;
  short: string;
  title: string;
  detail: string;
  tone: string;
  fields: RegisterField[];
};

export const registerModules: RegisterModule[] = [
  {
    id: "raw-material", short: "RM", title: "Raw material purchase", detail: "Supplier, vehicle, quantity and incoming quality", tone: "green",
    fields: [
      { key:"purchaseDate", label:"Purchase date", type:"date", group:"Purchase details", required:true },
      { key:"materialType", label:"Material type", type:"select", group:"Purchase details", required:true, options:["SMC","Cow dung","Cow dung slurry","Coco Peat","FOM","Press mud","Crop residue","Other"] },
      { key:"lotCode", label:"RM lot code — auto-generated", type:"text", group:"Purchase details", required:true, placeholder:"Select material type first" },
      { key:"supplier", label:"Supplier / dairy details", type:"text", group:"Purchase details", required:true },
      { key:"quantityKg", label:"Quantity", type:"number", group:"Purchase details", required:true, step:"0.01", unit:"kg" },
      { key:"ratePerUnit", label:"Rate per kg", type:"number", group:"Purchase details", required:true, step:"0.01", unit:"₹" },
      { key:"vehicleNo", label:"Vehicle number", type:"text", group:"Logistics & payment" },
      { key:"invoiceNo", label:"Vehicle / invoice number", type:"text", group:"Logistics & payment" },
      { key:"transportCost", label:"Transportation cost", type:"number", group:"Logistics & payment", step:"0.01", unit:"₹" },
      { key:"paymentStatus", label:"Payment status", type:"select", group:"Logistics & payment", options:["Pending","Part paid","Paid"] },
      { key:"moisturePct", label:"Moisture", type:"number", group:"Incoming quality", step:"0.01", unit:"%" },
      { key:"temperatureC", label:"Temperature", type:"number", group:"Incoming quality", step:"0.1", unit:"°C" },
      { key:"odorObservation", label:"Odour observation", type:"select", group:"Incoming quality", options:["No","Normal","Mild","Strong","Foul"] },
      { key:"contaminationCheck", label:"Contamination check", type:"select", group:"Incoming quality", options:["Yes","No","Clear","Minor contamination","Rejected"] },
      { key:"sampleCollectionDate", label:"Sample collection date", type:"date", group:"Sampling & lab" },
      { key:"sampleLabDate", label:"Sample sent to lab", type:"date", group:"Sampling & lab" },
      { key:"labStatus", label:"Lab report status", type:"select", group:"Sampling & lab", options:["Pending","Accepted","Not accepted"] },
      { key:"sealNo", label:"Seal number", type:"text", group:"Sampling & lab" },
      { key:"qualityRemarks", label:"Quality remarks", type:"textarea", group:"Sampling & lab" },
    ],
  },
  {
    id:"pre-composting", short:"PC", title:"Pre-composting", detail:"Decomposition event, turning, heat and moisture", tone:"blue",
    fields:[
      { key:"eventDate", label:"Entry date", type:"date", group:"Lot & event", required:true },
      { key:"lotCode", label:"Pre-compost lot code", type:"text", group:"Lot & event", required:true, placeholder:"PC/20260817/B04" },
      { key:"linkedRmLot", label:"Linked RM lot", type:"text", group:"Lot & event", required:true },
      { key:"eventType", label:"Event type", type:"select", group:"Lot & event", required:true, options:["Lot setup","Culture application","Turning","Watering","Temperature check","Transfer to bed"] },
      { key:"startDate", label:"Start date", type:"date", group:"Lot setup" },
      { key:"cultureApplicationDate", label:"Culture application date", type:"date", group:"Lot setup" },
      { key:"cultureDose", label:"Culture dose / preparation", type:"textarea", group:"Lot setup" },
      { key:"endDate", label:"End date / transfer", type:"date", group:"Lot setup" },
      { key:"totalWeightKg", label:"Total weight", type:"number", group:"Lot setup", step:"0.01", unit:"kg" },
      { key:"turningDate", label:"Turning date", type:"date", group:"Field reading" },
      { key:"avgTempC", label:"Average temperature", type:"number", group:"Field reading", step:"0.1", unit:"°C" },
      { key:"moisturePct", label:"Moisture maintained", type:"number", group:"Field reading", step:"0.1", unit:"%" },
      { key:"wateringDate", label:"Watering date", type:"date", group:"Field reading" },
      { key:"secondTurningDate", label:"Second turning date", type:"date", group:"Second field reading" },
      { key:"secondAvgTempC", label:"Second average temperature", type:"number", group:"Second field reading", step:"0.1", unit:"°C" },
      { key:"secondMoisturePct", label:"Second moisture reading", type:"number", group:"Second field reading", step:"0.1", unit:"%" },
      { key:"secondWateringDate", label:"Second watering date", type:"date", group:"Second field reading" },
      { key:"thirdTurningDate", label:"Third turning date", type:"date", group:"Third field reading" },
      { key:"thirdAvgTempC", label:"Third average temperature", type:"number", group:"Third field reading", step:"0.1", unit:"°C" },
      { key:"thirdMoisturePct", label:"Third moisture reading", type:"number", group:"Third field reading", step:"0.1", unit:"%" },
      { key:"thirdWateringDate", label:"Third watering date", type:"date", group:"Third field reading" },
      { key:"linkedBeds", label:"Linked production beds", type:"text", group:"Transfer & remarks", placeholder:"BED-01 to BED-20" },
      { key:"remarks", label:"Remarks", type:"textarea", group:"Transfer & remarks" },
    ],
  },
  {
    id:"bed-ops", short:"BD", title:"Bed production & operations", detail:"Bed lifecycle, worms, watering and status", tone:"plum",
    fields:[
      { key:"eventDate", label:"Entry date", type:"date", group:"Bed event", required:true },
      { key:"bedNumber", label:"Bed number", type:"text", group:"Bed event", required:true, placeholder:"BED-81" },
      { key:"eventType", label:"Operation", type:"select", group:"Bed event", required:true, options:["Full lifecycle import","Bed filling","Worm application","Feed application","Watering","Shuffling","Observation","Harvest update"] },
      { key:"currentStatus", label:"Current status", type:"select", group:"Bed event", required:true, options:["Planned","Filled","Inoculated","Active","Ready for harvest","Harvested","Hold"] },
      { key:"dimensions", label:"Bed dimensions", type:"text", group:"Setup details", placeholder:"12 x 4 x 2 ft" },
      { key:"fillDate", label:"Bed filling date", type:"date", group:"Setup details" },
      { key:"dungWeightKg", label:"Dung / substrate weight", type:"number", group:"Setup details", step:"0.01", unit:"kg" },
      { key:"wormSpecies", label:"Worm species & source", type:"text", group:"Worm & feed" },
      { key:"wormApplicationDate", label:"Earthworm application date", type:"date", group:"Worm & feed" },
      { key:"wormsAddedKg", label:"Earthworms added", type:"number", group:"Worm & feed", step:"0.01", unit:"kg" },
      { key:"foodApplicationDate", label:"Worm food application date", type:"date", group:"Worm & feed" },
      { key:"watering", label:"Watering frequency / quantity", type:"text", group:"Operations" },
      { key:"shufflingDate", label:"Shuffling date", type:"date", group:"Operations" },
      { key:"expectedHarvestDate", label:"Expected harvest date", type:"date", group:"Harvest linkage" },
      { key:"actualHarvestDate", label:"Actual harvest date", type:"date", group:"Harvest linkage" },
      { key:"productionBatchCode", label:"Production batch code", type:"text", group:"Harvest linkage", placeholder:"VB-26/B04" },
      { key:"linkedPreCompostLot", label:"Linked pre-compost lot", type:"text", group:"Harvest linkage" },
      { key:"harvestCompleteDate", label:"Harvest complete date", type:"date", group:"Harvest linkage" },
      { key:"harvestCompletionStatus", label:"Harvest completion status", type:"text", group:"Harvest linkage" },
      { key:"totalHarvestQtyKg", label:"Total harvest quantity", type:"number", group:"Harvest linkage", step:"0.01", unit:"kg" },
      { key:"remarks", label:"Supervisor remarks", type:"textarea", group:"Harvest linkage" },
    ],
  },
  {
    id:"harvest", short:"HV", title:"Harvesting & packaging", detail:"Net production, packaging and batch traceability", tone:"orange",
    fields:[
      { key:"harvestDate", label:"Harvest date", type:"date", group:"Harvest details", required:true },
      { key:"bedNumber", label:"Bed number", type:"text", group:"Harvest details", required:true },
      { key:"netYieldKg", label:"Net yield after sieving", type:"number", group:"Harvest details", required:true, step:"0.01", unit:"kg" },
      { key:"productionBatchCode", label:"Production batch linked — auto-generated", type:"text", group:"Batch & packing", required:true },
      { key:"batchNumber", label:"Finished batch number", type:"text", group:"Batch & packing", required:true, placeholder:"VCM-26/B04" },
      { key:"packagingType", label:"Packaging type", type:"select", group:"Batch & packing", options:["Bulk/Loose","50 Kg Bag","Bulk / loose","50 kg bag","25 kg bag","5 kg bag","1 kg pack","Other"] },
      { key:"packetsPrepared", label:"Packets prepared", type:"number", group:"Batch & packing", step:"1", unit:"pcs" },
      { key:"packagedQuantityKg", label:"Packaged output quantity", type:"number", group:"Batch & packing", required:true, step:"0.01", unit:"kg" },
      { key:"moisturePct", label:"Moisture content", type:"number", group:"Quality & tagging", step:"0.01", unit:"%" },
      { key:"tagging", label:"Tagging / label details", type:"textarea", group:"Quality & tagging" },
    ],
  },
  {
    id:"quality", short:"QC", title:"Quality control", detail:"Lab report, nutrients and FCO compliance", tone:"teal",
    fields:[
      { key:"samplingDate", label:"Sampling date", type:"date", group:"Sample & report", required:true },
      { key:"sealNo", label:"Sample seal number", type:"text", group:"Sample & report", required:true },
      { key:"labNameDate", label:"Lab name / sample sent date", type:"text", group:"Sample & report", required:true },
      { key:"productType", label:"Product type", type:"select", group:"Sample & report", required:true, options:["RM","FG"] },
      { key:"batchNumber", label:"Batch number", type:"text", group:"Sample & report", required:true },
      { key:"reportDate", label:"Test report date", type:"date", group:"Sample & report" },
      { key:"reportNo", label:"Lab test report number", type:"text", group:"Sample & report" },
      { key:"fcoStatus", label:"FCO compliance status", type:"select", group:"Sample & report", required:true, options:["Pending","Pass (FCO Compliant)","Conditional","Fail / Not compliant"] },
      { key:"ph", label:"pH", type:"number", group:"Core analysis", step:"0.01" },
      { key:"ec", label:"EC", type:"number", group:"Core analysis", step:"0.01" },
      { key:"moisturePct", label:"Moisture", type:"number", group:"Core analysis", step:"0.01", unit:"%" },
      { key:"organicCarbonPct", label:"Organic carbon", type:"number", group:"Core analysis", step:"0.01", unit:"%" },
      { key:"organicMatterPct", label:"Organic matter", type:"number", group:"Core analysis", step:"0.01", unit:"%" },
      { key:"cnRatio", label:"C/N ratio", type:"number", group:"Core analysis", step:"0.01" },
      { key:"nitrogenPct", label:"Nitrogen", type:"number", group:"Nutrients", step:"0.01", unit:"%" },
      { key:"phosphorusPct", label:"Phosphorous", type:"number", group:"Nutrients", step:"0.01", unit:"%" },
      { key:"potassiumPct", label:"Potassium", type:"number", group:"Nutrients", step:"0.01", unit:"%" },
      { key:"ashPct", label:"Ash", type:"number", group:"Nutrients", step:"0.01", unit:"%" },
      { key:"calciumPct", label:"Calcium", type:"number", group:"Nutrients", step:"0.01", unit:"%" },
      { key:"manganese", label:"Manganese", type:"number", group:"Micronutrients", step:"0.01" },
      { key:"iron", label:"Iron", type:"number", group:"Micronutrients", step:"0.01" },
      { key:"manganeseAdditional", label:"Manganese (second workbook reading)", type:"number", group:"Micronutrients", step:"0.01" },
      { key:"zinc", label:"Zinc", type:"number", group:"Micronutrients", step:"0.01" },
      { key:"copper", label:"Copper", type:"number", group:"Micronutrients", step:"0.01" },
    ],
  },
  {
    id:"sales", short:"SL", title:"Sales register", detail:"Invoice, customer, batch and dispatch revenue", tone:"gold",
    fields:[
      { key:"salesDate", label:"Sales date", type:"date", group:"Invoice & customer", required:true },
      { key:"invoiceNo", label:"Invoice number", type:"text", group:"Invoice & customer", required:true },
      { key:"customer", label:"Customer / dealer details", type:"text", group:"Invoice & customer", required:true },
      { key:"dealerDrcNo", label:"Dealer DRC number", type:"text", group:"Invoice & customer" },
      { key:"batchNumber", label:"Harvested batch number", type:"select", group:"Dispatch & value", required:true },
      { key:"quantityKg", label:"Quantity sold", type:"number", group:"Dispatch & value", required:true, step:"0.01", unit:"kg" },
      { key:"pricePerKg", label:"Price per kg", type:"number", group:"Dispatch & value", required:true, step:"0.01", unit:"₹" },
    ],
  },
  {
    id:"inventory", short:"ST", title:"Stock inventory", detail:"Production, outflow, loss and physical audit", tone:"forest",
    fields:[
      { key:"entryDate", label:"Entry date", type:"date", group:"Stock movement", required:true },
      { key:"productDescription", label:"Product description", type:"text", group:"Stock movement", required:true, placeholder:"Premium Organic Vermicompost" },
      { key:"packagingSize", label:"Packaging size", type:"select", group:"Stock movement", required:true, options:["Bulk / Loose","Bulk / loose","Bags","50 kg bag","25 kg bag","5 kg bag","1 kg pack","Mixed bags"] },
      { key:"openingStockKg", label:"Opening stock", type:"number", group:"Stock movement", step:"0.01", unit:"kg" },
      { key:"productionKg", label:"Production (+)", type:"number", group:"Stock movement", step:"0.01", unit:"kg" },
      { key:"salesOutflowKg", label:"Sales outflow (-)", type:"number", group:"Stock movement", step:"0.01", unit:"kg" },
      { key:"damageLossKg", label:"Damage / loss (-)", type:"number", group:"Stock movement", step:"0.01", unit:"kg" },
      { key:"auditStatus", label:"Physical audit status", type:"select", group:"Audit", required:true, options:["Pending","Verified (Match)","Shortage","Excess","Under review"] },
      { key:"remarks", label:"Audit remarks", type:"textarea", group:"Audit" },
      { key:"sourceDateNote", label:"Source date note", type:"text", group:"Audit" },
    ],
  },
  {
    id:"expenses", short:"EX", title:"Cost & expenses", detail:"Expense category, vendor, payment and optimization", tone:"red",
    fields:[
      { key:"expenseDate", label:"Expense date", type:"date", group:"Expense details", required:true },
      { key:"category", label:"Expense category", type:"select", group:"Expense details", required:true, options:["Raw materials","Logistics","Labor wages","Packaging cost","Compliance & government","Culture","Earthworm food","Utilities","Maintenance","Other"] },
      { key:"description", label:"Description / vendor details", type:"textarea", group:"Expense details", required:true },
      { key:"paymentMode", label:"Payment mode", type:"select", group:"Payment & action", required:true, options:["Bank Transfer","Bank transfer","Online","Cash","Cheque","Credit"] },
      { key:"expenseCost", label:"Expense cost", type:"number", group:"Payment & action", required:true, step:"0.01", unit:"₹" },
      { key:"optimizationLog", label:"Cost optimization / compliance log", type:"textarea", group:"Payment & action" },
      { key:"sourceCategory", label:"Workbook expense category", type:"text", group:"Payment & action" },
      { key:"sourceDateNote", label:"Source date note", type:"text", group:"Payment & action" },
    ],
  },
];

export const registerModuleMap = Object.fromEntries(registerModules.map((module) => [module.id, module]));
