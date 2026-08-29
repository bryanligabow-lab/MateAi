export type Language = "es" | "en";

export const serviceCatalog = [
  { slug:"ai-agents", icon:"AI", es:{name:"Agentes de IA",description:"Agentes inteligentes que atienden, califican, coordinan tareas y trabajan junto a tu equipo."}, en:{name:"AI Agents",description:"Intelligent agents that assist, qualify, coordinate tasks, and work alongside your team."} },
  { slug:"web-chatbots", icon:"01", featured:true, es:{name:"Chatbots para páginas web",description:"Convierte cada visita en una conversación, un contacto calificado o una cita."}, en:{name:"Website Chatbots",description:"Turn every website visit into a conversation, a qualified lead, or an appointment."} },
  { slug:"social-media", icon:"@", es:{name:"Redes sociales",description:"Respuestas, seguimiento y captación automatizada para tus canales sociales."}, en:{name:"Social Media",description:"Automated replies, follow-up, and lead capture across your social channels."} },
  { slug:"websites", icon:"//", es:{name:"Creación de páginas web",description:"Sitios modernos, rápidos y diseñados para convertir visitantes en clientes."}, en:{name:"Website Creation",description:"Modern, fast websites designed to turn visitors into customers."} },
  { slug:"custom-systems", icon:"<>", es:{name:"Sistemas a medida",description:"Software construido alrededor de los procesos y necesidades reales de tu empresa."}, en:{name:"Custom Systems",description:"Software built around your company’s real processes and requirements."} },
  { slug:"custom-automation", icon:"∞", es:{name:"Automatización y agentes personalizados",description:"Automatizamos cualquier flujo y creamos el agente de IA que tu operación necesita."}, en:{name:"Custom Automation & Agents",description:"We automate any workflow and create the AI agent your operation needs."} },
] as const;

export type Service = (typeof serviceCatalog)[number];
export function getService(slug:string){return serviceCatalog.find(service=>service.slug===slug)}
