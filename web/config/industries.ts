export type IndustryKey = "roofing" | "hvac" | "plumbing" | "junk-removal" | "cleaning" | "landscaping" | "auto-services";

export type Industry = {
  key: IndustryKey;
  name: string;
  headline: string;
  description: string;
  painPoints: string[];
  conversation: { role: "visitor" | "ai"; text: string }[];
};

export const industries: Industry[] = [
  { key:"roofing", name:"Roofing", headline:"AI Receptionist for Roofing Companies", description:"Capture roofing leads even when nobody is available to answer.", painPoints:["Storm leads arrive after hours","Inspection requests need fast follow-up","Every unanswered call risks a competitor"], conversation:[{role:"visitor",text:"We had hail last night. Can someone inspect the roof?"},{role:"ai",text:"Absolutely. What ZIP code is the property in?"},{role:"visitor",text:"84043"},{role:"ai",text:"Thanks. I can help request an inspection. What's the best phone number for the team to reach you?"}] },
  { key:"hvac", name:"HVAC", headline:"AI Receptionist for HVAC Companies", description:"Turn urgent heating and cooling questions into qualified service calls.", painPoints:["No-cool calls cannot wait","Seasonal spikes overwhelm teams","Dispatch needs clear job details"], conversation:[{role:"visitor",text:"Our AC stopped cooling. Do you have availability today?"},{role:"ai",text:"I can help with that. Is the system running but blowing warm air, or not turning on?"}] },
  { key:"plumbing", name:"Plumbing", headline:"AI Receptionist for Plumbing Companies", description:"Answer urgent plumbing questions and capture the details your dispatcher needs.", painPoints:["Emergency intent needs instant response","Incomplete information slows dispatch","After-hours leads call the next company"], conversation:[{role:"visitor",text:"Do you offer emergency plumbing?"},{role:"ai",text:"Yes. We provide 24/7 emergency plumbing service. What's your ZIP code?"},{role:"visitor",text:"84043"},{role:"ai",text:"Great. What's the best phone number for our team to reach you?"}] },
  { key:"junk-removal", name:"Junk Removal", headline:"AI Receptionist for Junk Removal Companies", description:"Qualify pickup requests and move visitors toward an estimate.", painPoints:["Customers expect quick estimates","Photos and item details matter","Fast responses win local jobs"], conversation:[{role:"visitor",text:"Can you remove a couch and old refrigerator?"},{role:"ai",text:"Yes. What city is the pickup in, and are either items upstairs?"}] },
  { key:"cleaning", name:"Cleaning", headline:"AI Receptionist for Cleaning Companies", description:"Capture home details, frequency, and booking intent around the clock.", painPoints:["Quote details are repetitive","Recurring service requires qualification","Teams cannot answer while cleaning"], conversation:[{role:"visitor",text:"How much for a biweekly house cleaning?"},{role:"ai",text:"I can help estimate it. How many bedrooms and bathrooms does the home have?"}] },
  { key:"landscaping", name:"Landscaping", headline:"AI Receptionist for Landscaping Companies", description:"Qualify projects, service areas, and estimate requests before your team calls back.", painPoints:["Project scope varies widely","Seasonal demand creates backlogs","Service-area checks waste phone time"], conversation:[{role:"visitor",text:"Do you install sprinklers in Lehi?"},{role:"ai",text:"Yes, that area is covered. Is this a new installation or a repair?"}] },
  { key:"auto-services", name:"Auto Services", headline:"AI Receptionist for Auto Service Businesses", description:"Answer service questions and capture appointment requests while technicians stay focused.", painPoints:["Technicians cannot stop for every call","Customers need service clarity","Missed calls become missed appointments"], conversation:[{role:"visitor",text:"Can you look at my brakes this week?"},{role:"ai",text:"I can help request an appointment. What vehicle year, make, and model do you have?"}] },
];

export const demoIndustries = industries.filter((item) => ["roofing","hvac","plumbing","junk-removal"].includes(item.key));
export function getIndustry(key: string) { return industries.find((industry) => industry.key === key); }
