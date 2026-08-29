export type PaddleEvent = { event_id:string; event_type:string; occurred_at:string; data:{ id:string; customer_id?:string; subscription_id?:string; status?:string; custom_data?:Record<string,unknown>; items?:Array<{price?:{id?:string}}> } };
const handledEvents=new Set(["transaction.completed","transaction.payment_failed","subscription.created","subscription.updated","subscription.canceled","subscription.paused","subscription.resumed"]);

export async function processPaymentEvent(db:D1Database,event:PaddleEvent){
 if(!handledEvents.has(event.event_type))return {handled:false};
 const existing=await db.prepare("SELECT event_id FROM webhook_events WHERE event_id = ?").bind(event.event_id).first();
 if(existing)return {handled:true,duplicate:true};
 const data=event.data;const subscriptionId=event.event_type.startsWith("subscription.")?data.id:data.subscription_id||null;const plan=String(data.custom_data?.plan||"unknown");
 const statements=[db.prepare("INSERT INTO webhook_events (event_id,event_type,occurred_at,processed_at) VALUES (?,?,?,datetime('now'))").bind(event.event_id,event.event_type,event.occurred_at)];
 if(subscriptionId)statements.push(db.prepare("INSERT INTO subscriptions (paddle_subscription_id,paddle_customer_id,plan,status,updated_at) VALUES (?,?,?,?,datetime('now')) ON CONFLICT(paddle_subscription_id) DO UPDATE SET paddle_customer_id=excluded.paddle_customer_id,plan=excluded.plan,status=excluded.status,updated_at=datetime('now')").bind(subscriptionId,data.customer_id||null,plan,data.status||event.event_type.split(".")[1]));
 await db.batch(statements);return {handled:true,duplicate:false};
}

export async function requestProvisioning(db:D1Database,event:PaddleEvent){
 if(event.event_type!=="transaction.completed")return;
 await db.prepare("INSERT OR IGNORE INTO provisioning_jobs (event_id,transaction_id,status,created_at) VALUES (?,?,'pending',datetime('now'))").bind(event.event_id,event.data.id).run();
}
