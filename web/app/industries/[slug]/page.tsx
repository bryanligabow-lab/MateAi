import { notFound } from "next/navigation";import type { Metadata } from "next";import { industries,getIndustry } from "@/config/industries";import { IndustryPage } from "@/components/marketing/IndustryPage";
export function generateStaticParams(){return industries.map(i=>({slug:i.key}))}
export async function generateMetadata({params}:{params:Promise<{slug:string}>}):Promise<Metadata>{const {slug}=await params;const i=getIndustry(slug);return i?{title:i.headline,description:i.description,alternates:{canonical:`/${i.key}`}}:{}}
export default async function Page({params}:{params:Promise<{slug:string}>}){const {slug}=await params;const i=getIndustry(slug);if(!i)notFound();return <IndustryPage industry={i}/>}
