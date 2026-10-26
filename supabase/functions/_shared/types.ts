export type ProviderKind = "microsoft" | "microsoft_app" | "gmail" | "imap";

export type Category = "invoice" | "utility_bill" | "vendor_info" | "statement" | "other";

export type ItemStatus = "done" | "action_needed" | "flagged" | "info" | "skipped";

export type ReconStatus = "found_in_ledger" | "received_by_email" | "amount_mismatch" | "missing";

export interface Mailbox {
  id: string;
  owner_user_id: string;
  provider: ProviderKind;
  email_address: string;
  display_name: string | null;
  credentials_secret_id: string | null;
  smart_ap_address: string | null;
  auto_forward_invoices: boolean;
  lookback_hours: number;
  report_recipient: string | null;
  processed_label: string;
  portal_vendors: { name: string; portal: string }[];
  managed_entities: string[];
  active: boolean;
  last_run_at: string | null;
}

/** Lightweight listing entry, enough to dedupe and order before fetching full content. */
export interface MailSummary {
  id: string;
  threadId: string | null;
  receivedAt: Date;
  labels: string[];
}

export interface AttachmentMeta {
  id: string;
  name: string;
  contentType: string;
  size: number;
  /** Some providers return content inline with the message. */
  content?: Uint8Array;
}

export interface MailMessage {
  id: string;
  threadId: string | null;
  from: { address: string; name: string | null };
  replyTo: string[];
  to: string[];
  subject: string;
  receivedAt: Date;
  bodyText: string;
  /** Raw header values we care about for the safety check. */
  headers: Record<string, string>;
  attachments: AttachmentMeta[];
  labels: string[];
}

export interface OutgoingAttachment {
  name: string;
  contentType: string;
  content: Uint8Array;
}

export interface MailProvider {
  /** Inbox messages received since `since`, excluding mail the owner sent. */
  listMessages(since: Date): Promise<MailSummary[]>;
  getMessage(id: string): Promise<MailMessage>;
  getAttachment(messageId: string, attachment: AttachmentMeta): Promise<Uint8Array>;
  /** Forward the original message (with its attachments) to `to`, with a short comment on top. */
  forward(message: MailMessage, to: string, comment: string): Promise<void>;
  sendHtml(to: string, subject: string, html: string): Promise<void>;
  /** Tag the message with a category / label / keyword so it is visibly processed. */
  addLabel(messageId: string, label: string): Promise<void>;
  close?(): Promise<void>;
}

// ---------------------------------------------------------------------------
// Classifier output (mirrors the JSON schema in classify.ts)
// ---------------------------------------------------------------------------

export interface InvoiceFields {
  invoice_number: string | null;
  invoice_date: string | null;
  amount: number | null;
  currency: string | null;
  due_date: string | null;
  po_number: string | null;
  billed_entity: string | null;
}

export interface UtilityFields {
  provider: string | null;
  account_numbers: string[];
  billing_period: string | null;
  amount: number | null;
  due_date: string | null;
  portal_url: string | null;
}

export interface StatementLine {
  invoice_number: string | null;
  invoice_date: string | null;
  amount: number | null;
  balance: number | null;
  status: string | null;
}

export interface StatementFields {
  statement_date: string | null;
  closing_balance: number | null;
  lines: StatementLine[];
}

export interface Classification {
  category: Category;
  other_subtype:
    | "past_due"
    | "dispute_or_credit"
    | "payment_confirmation"
    | "bank_detail_change"
    | "not_ap"
    | "unclear"
    | null;
  confidence: "high" | "medium" | "low";
  reasoning: string;
  vendor_name: string | null;
  one_line_summary: string;
  invoice: InvoiceFields | null;
  utility: UtilityFields | null;
  vendor_info: { summary: string; key_dates: string[] } | null;
  statement: StatementFields | null;
  fraud_risk: "none" | "low" | "high";
  fraud_signals: string[];
  urgent: boolean;
  employee_action: string;
}
