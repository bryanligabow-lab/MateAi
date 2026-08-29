import type { Metadata } from "next";
import { Hero } from "@/components/marketing/Hero";
import { BeforeAfter, Capabilities, DashboardPreview, FinalCta, HowItWorks, Integrations, ProblemSection, RoiCalculator, StorySection } from "@/components/marketing/HomeSections";
export const metadata:Metadata={title:"Website Chatbots",description:"AI website chatbots that answer questions, capture leads, and book appointments."};
export default function WebChatbots(){return <><Hero/><div className="proof-strip"><span>Built for local service businesses.</span><span>Designed to capture customers when your team can&apos;t.</span></div><ProblemSection/><StorySection/><Capabilities/><HowItWorks/><BeforeAfter/><DashboardPreview/><Integrations/><RoiCalculator/><FinalCta/></>}
