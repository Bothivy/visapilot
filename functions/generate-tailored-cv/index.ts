import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import { extractText, getDocumentProxy } from "npm:unpdf@1.8.1";
import { AlignmentType, BorderStyle, Document, HeadingLevel, Packer, Paragraph, TextRun } from "npm:docx@9.5.1";
import { PDFDocument, StandardFonts, rgb } from "npm:pdf-lib@1.17.1";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};
const json = (body: unknown, status = 200) => new Response(JSON.stringify(body), {
  status,
  headers: { ...corsHeaders, "Content-Type": "application/json" },
});
const stopWords = new Set("the and for with from this that your you are our their they them job role work working will would can required essential desirable experience skills skill ability team teams about using use who what when where how not but all any have has had into position candidate successful including within across through also such must should duties responsibilities person employer organisation provide providing support supporting".split(" "));

function clean(value: unknown) {
  return String(value || "").replace(/<[^>]*>/g, " ").replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&").replace(/\s+/g, " ").trim();
}
function words(value: unknown) {
  return clean(value).toLowerCase().match(/[a-z][a-z0-9+#.-]{2,}/g) || [];
}
function importantWords(value: unknown, limit = 20) {
  const frequency = new Map<string, number>();
  for (const word of words(value)) {
    if (stopWords.has(word)) continue;
    frequency.set(word, (frequency.get(word) || 0) + 1);
  }
  return [...frequency.entries()].sort((a, b) => b[1] - a[1]).slice(0, limit).map(([word]) => word);
}
function planFromPrice(priceId: string) {
  const map: Record<string, string> = {
    price_1U8NJv2MpjlYziXhFvFSUmmf: "copilot_monthly",
    price_1U8Ng32MpjlYziXho9C7LOUU: "copilot_quarterly",
    price_1U8NjP2MpjlYziXhw7aiSV15: "pro_monthly",
    price_1U8NkI2MpjlYziXhpecurry4: "pro_quarterly",
    price_1U8NlP2MpjlYziXh6Bt7MPBY: "copilot_pass_7d",
  };
  return map[priceId] || "";
}
function normaliseDocumentText(value: unknown) {
  const source = Array.isArray(value) ? value.join("\n") : String(value || "");
  return source.replace(/<[^>]*>/g, " ").replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&").replace(/\r/g, "\n")
    .split(/\n+/).map((line) => line.replace(/\s+/g, " ").trim()).filter(Boolean).join("\n").trim();
}
async function readableCvText(admin: any, cv: any) {
  let extracted = "";
  if (cv?.storage_bucket && cv?.storage_path) {
    const { data: file, error } = await admin.storage.from(cv.storage_bucket).download(cv.storage_path);
    if (!error && file) {
      const filename = String(cv?.original_filename || cv?.storage_path || "").toLowerCase();
      if (filename.endsWith(".pdf")) {
        const pdf = await getDocumentProxy(new Uint8Array(await file.arrayBuffer()));
        extracted = normaliseDocumentText((await extractText(pdf, { mergePages: true }))?.text);
      } else if (filename.endsWith(".txt")) extracted = normaliseDocumentText(await file.text());
      else if (filename.endsWith(".rtf")) extracted = normaliseDocumentText((await file.text()).replace(/\\'[0-9a-f]{2}/gi, " ").replace(/\\[a-z]+-?\d* ?/gi, " ").replace(/[{}]/g, " "));
    }
  }
  if (extracted.length < 80) extracted = normaliseDocumentText(cv?.cv_text);
  if (extracted.length >= 80 && extracted !== cv?.cv_text) {
    await admin.from("candidate_cvs").update({ cv_text: extracted.slice(0, 50000), updated_at: new Date().toISOString() }).eq("id", cv.id);
  }
  return extracted;
}
type CvBlock = { heading: string; bullets: string[] };
type CvSections = { summary: string; skills: string[]; experience: CvBlock[]; education: string[]; qualifications: string[]; languages: string[] };
const sectionAliases: Record<string, keyof Omit<CvSections, "summary"> | "summary"> = {
  "summary": "summary", "profile": "summary", "professional profile": "summary", "personal profile": "summary",
  "experience": "experience", "professional experience": "experience", "work experience": "experience", "employment history": "experience", "career history": "experience",
  "experience, education and qualifications": "experience",
  "skills": "skills", "key skills": "skills", "core skills": "skills", "competencies": "skills",
  "education": "education", "education and training": "education", "academic background": "education",
  "qualifications": "qualifications", "certifications": "qualifications", "certificates": "qualifications", "training": "qualifications",
  "languages": "languages",
};
const actionStart = /\b(assisted|provided|supported|managed|maintained|collaborated|monitored|delivered|developed|addressed|processed|analysed|analyzed|engaged|established|aided|changed|promoted|led|coordinated|implemented|resolved|created|improved|worked|performed|ensured|prepared|administered|recorded|communicated|handled|organised|organized|trained|achieved)\b/i;
function cleanItem(value: string) { return value.replace(/^[-•▪◦]\s*/, "").replace(/\s+/g, " ").trim(); }
function splitItems(value: string) {
  return value.split(/\n+|(?<=[.!?])\s+(?=[A-Z])/).map(cleanItem).filter((item) => item.length > 2);
}
function experienceBlocks(value: string): CvBlock[] {
  const blocks: CvBlock[] = [];
  let current: CvBlock | null = null;
  for (const item of splitItems(value)) {
    const action = item.search(actionStart);
    if (action > 2 && action < 120) {
      const possibleHeading = cleanItem(item.slice(0, action));
      const bullet = cleanItem(item.slice(action));
      if (possibleHeading.length <= 110) {
        current = { heading: possibleHeading, bullets: [] };
        blocks.push(current);
        if (bullet) current.bullets.push(bullet);
        continue;
      }
    }
    if (action === 0 || item.length > 110) {
      if (!current) { current = { heading: "Relevant experience", bullets: [] }; blocks.push(current); }
      current.bullets.push(item);
    } else {
      current = { heading: item, bullets: [] };
      blocks.push(current);
    }
  }
  return blocks.filter((block) => block.heading || block.bullets.length).slice(0, 12);
}
function parseCvSections(text: string): CvSections {
  let source = normaliseDocumentText(text);
  // Older VisaPilot exports placed the original CV beneath one combined heading.
  // Recover the title-cased headings embedded in that flattened legacy text.
  source = source.replace(/\b(Summary|Experience|Skills|Languages|Education|Qualifications)\b/g, "\n$1\n");
  const known = Object.keys(sectionAliases).sort((a, b) => b.length - a.length).map((value) => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|");
  if (!source.split("\n").some((line) => sectionAliases[line.toLowerCase()])) {
    source = source.replace(new RegExp(`\\b(${known})\\b`, "gi"), "\n$1\n");
  }
  const buckets: Record<string, string[]> = { summary: [], experience: [], skills: [], education: [], qualifications: [], languages: [], prelude: [] };
  let current = "prelude";
  for (const line of source.split(/\n+/).map(cleanItem).filter(Boolean)) {
    const key = sectionAliases[line.toLowerCase()];
    if (key) { current = key; continue; }
    buckets[current].push(line);
  }
  const summary = splitItems(buckets.summary.join("\n")).slice(0, 3).join(" ");
  const skillCatalog = ["strong communication skills", "communication", "patience and calm", "cleaning and sanitation", "work prioritization", "team work", "teamwork", "personal care", "patient record keeping", "manual handling", "health monitoring", "confidentiality", "patient handling and positioning", "food hygiene", "patient safeguarding", "health and safety", "bedmaking", "bathing and dressing assistance", "customer service", "complaint resolution", "record keeping"];
  const skillSource = buckets.skills.join(" ").toLowerCase();
  const explicitSkills = buckets.skills.flatMap((line) => line.split(/[,;|•▪◦]+/)).map(cleanItem).filter((item) => item.length > 1 && item.length < 70);
  const detectedSkills = skillCatalog.filter((skill) => skillSource.includes(skill));
  const skills = [...new Set([...explicitSkills, ...detectedSkills])];
  const education = splitItems(buckets.education.join("\n"));
  const qualifications = splitItems(buckets.qualifications.join("\n"));
  const languageSource = `${buckets.languages.join(" ")} ${source}`.toLowerCase();
  const languages = ["English", "Welsh", "Nepali", "Hindi", "Urdu", "Punjabi", "Bengali", "French", "Spanish", "Arabic"].filter((language) => new RegExp(`\\b${language}\\b`, "i").test(languageSource));
  let experienceText = buckets.experience.join("\n");
  if (!experienceText && buckets.prelude.length) experienceText = buckets.prelude.join("\n");
  const experience = experienceBlocks(experienceText);
  const dates = source.match(/\b(?:0?[1-9]|1[0-2])\/\d{4}(?:\s+to\s+(?:Current|(?:0?[1-9]|1[0-2])\/\d{4}))?\b/gi) || [];
  experience.forEach((block, index) => { if (dates[index] && !block.heading.includes(dates[index])) block.heading = `${block.heading} · ${dates[index]}`; });
  return { summary, skills, experience, education, qualifications, languages };
}
function wrapText(text: string, font: any, size: number, width: number) {
  const lines: string[] = [];
  for (const paragraph of String(text).split(/\n/)) {
    if (!paragraph.trim()) { lines.push(""); continue; }
    let line = "";
    for (const word of paragraph.trim().split(/\s+/)) {
      const candidate = line ? `${line} ${word}` : word;
      if (font.widthOfTextAtSize(candidate, size) <= width) line = candidate;
      else { if (line) lines.push(line); line = word; }
    }
    if (line) lines.push(line);
  }
  return lines;
}
async function cvDocx(name: string, contact: string, title: string, sections: CvSections) {
  const sectionHeading = (text: string) => new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 220, after: 90 },
    border: { bottom: { color: "D9D2EA", style: BorderStyle.SINGLE, size: 5, space: 4 } },
    children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 22, font: "Arial", color: "4F2AB8" })],
  });
  const bullet = (text: string) => new Paragraph({ bullet: { level: 0 }, spacing: { after: 70, line: 270 }, children: [new TextRun({ text, size: 20, font: "Arial", color: "25212D" })] });
  const children: Paragraph[] = [
    new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 45 }, children: [new TextRun({ text: name, bold: true, size: 38, font: "Arial", color: "1F1637" })] }),
    new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 45 }, children: [new TextRun({ text: title, bold: true, size: 23, font: "Arial", color: "4F2AB8" })] }),
    new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 170 }, children: [new TextRun({ text: contact, size: 19, font: "Arial", color: "5E5868" })] }),
    sectionHeading("Professional profile"),
    new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 120, line: 285 }, children: [new TextRun({ text: sections.summary, size: 20, font: "Arial", color: "25212D" })] }),
    sectionHeading("Core skills"),
    ...sections.skills.map(bullet),
    sectionHeading("Professional experience"),
  ];
  for (const block of sections.experience) {
    children.push(new Paragraph({ keepNext: true, spacing: { before: 100, after: 55 }, children: [new TextRun({ text: block.heading, bold: true, size: 21, font: "Arial", color: "1F1637" })] }));
    children.push(...block.bullets.map(bullet));
  }
  if (sections.education.length) children.push(sectionHeading("Education"), ...sections.education.map(bullet));
  if (sections.qualifications.length) children.push(sectionHeading("Qualifications and training"), ...sections.qualifications.map(bullet));
  if (sections.languages.length) children.push(sectionHeading("Languages"), ...sections.languages.map(bullet));
  const document = new Document({ sections: [{
    properties: { page: { margin: { top: 780, right: 850, bottom: 780, left: 850 } } },
    children,
  }] });
  return new Uint8Array(await Packer.toBuffer(document));
}
async function cvPdf(name: string, contact: string, title: string, sections: CvSections) {
  const pdf = await PDFDocument.create();
  const regular = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const margin = 54, width = 595.28, height = 841.89;
  let page = pdf.addPage([width, height]), y = height - margin;
  const line = (text: string, font: any, size: number, gap: number, colour = rgb(0.15, 0.13, 0.18), indent = 0) => {
    for (const value of wrapText(text, font, size, width - margin * 2)) {
      if (y < margin + 25) { page = pdf.addPage([width, height]); y = height - margin; }
      if (value) page.drawText(value, { x: margin + indent, y, size, font, color: colour });
      y -= gap;
    }
  };
  const heading = (text: string) => {
    if (y < margin + 50) { page = pdf.addPage([width, height]); y = height - margin; }
    y -= 7; line(text.toUpperCase(), bold, 10.5, 15, rgb(0.31, 0.16, 0.72));
    page.drawLine({ start: { x: margin, y: y + 7 }, end: { x: width - margin, y: y + 7 }, thickness: 0.7, color: rgb(0.84, 0.81, 0.91) });
    y -= 4;
  };
  const bullet = (text: string) => line(`- ${text}`, regular, 9.5, 14.5, rgb(0.15, 0.13, 0.18), 8);
  line(name, bold, 20, 24, rgb(0.12, 0.09, 0.22));
  line(title, bold, 11.5, 17, rgb(0.31, 0.16, 0.72));
  line(contact, regular, 9.5, 20, rgb(0.35, 0.32, 0.39));
  heading("Professional profile"); line(sections.summary, regular, 10, 15);
  heading("Core skills"); sections.skills.forEach(bullet);
  heading("Professional experience");
  for (const block of sections.experience) { y -= 3; line(block.heading, bold, 10.2, 15, rgb(0.12, 0.09, 0.22)); block.bullets.forEach(bullet); }
  if (sections.education.length) { heading("Education"); sections.education.forEach(bullet); }
  if (sections.qualifications.length) { heading("Qualifications and training"); sections.qualifications.forEach(bullet); }
  if (sections.languages.length) { heading("Languages"); sections.languages.forEach(bullet); }
  return new Uint8Array(await pdf.save());
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  if (req.method !== "POST") return json({ error: "Method not allowed." }, 405);
  try {
    const supabaseUrl = Deno.env.get("SUPABASE_URL"), anonKey = Deno.env.get("SUPABASE_ANON_KEY"), serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
    if (!supabaseUrl || !anonKey || !serviceKey) throw new Error("Missing required environment variable.");
    const authorization = req.headers.get("Authorization") || "";
    if (!authorization.startsWith("Bearer ")) return json({ error: "Unauthorized" }, 401);
    const candidate = createClient(supabaseUrl, anonKey, { global: { headers: { Authorization: authorization } }, auth: { persistSession: false, autoRefreshToken: false } });
    const { data: { user }, error: userError } = await candidate.auth.getUser();
    if (userError || !user) return json({ error: "Unauthorized" }, 401);
    const admin = createClient(supabaseUrl, serviceKey, { auth: { persistSession: false, autoRefreshToken: false } });
    const body = await req.json().catch(() => ({}));
    const savedJobId = String(body?.saved_job_id || "");
    if (!savedJobId) return json({ error: "A saved application is required." }, 400);
    const { data: billing, error: billingError } = await admin.from("billing_subscriptions").select("status,current_period_end,plan_key,stripe_price_id").eq("user_id", user.id).maybeSingle();
    if (billingError) throw billingError;
    const planKey = String(billing?.plan_key || planFromPrice(String(billing?.stripe_price_id || "")));
    const end = billing?.current_period_end ? new Date(billing.current_period_end).getTime() : null;
    const activePlan = ["active", "trialing"].includes(String(billing?.status || "")) && (!end || Number.isNaN(end) || end > Date.now());
    const allowanceLimit = activePlan && /^pro_/.test(planKey) ? -1 : activePlan && /^copilot_/.test(planKey) ? 25 : 5;
    const { data: saved, error: savedError } = await admin.from("saved_jobs").select("id,job_id,user_id").eq("id", savedJobId).eq("user_id", user.id).maybeSingle();
    if (savedError) throw savedError;
    if (!saved) return json({ error: "Saved application not found." }, 404);
    const [{ data: job, error: jobError }, { data: cv, error: cvError }, { data: profile, error: profileError }] = await Promise.all([
      admin.from("jobs").select("id,title,location,description,employer_id,employers(name)").eq("id", saved.job_id).maybeSingle(),
      admin.from("candidate_cvs").select("id,cv_name,original_filename,cv_text,storage_bucket,storage_path,updated_at").eq("user_id", user.id).order("updated_at", { ascending: false }).limit(1).maybeSingle(),
      admin.from("candidate_profiles").select("full_name,skills,experience_summary,target_roles").eq("user_id", user.id).maybeSingle(),
    ]);
    if (jobError) throw jobError; if (cvError) throw cvError; if (profileError) throw profileError;
    if (!job) return json({ error: "Job details could not be found." }, 404);
    if (!cv) return json({ error: "Upload your CV before requesting a review." }, 400);
    const cvText = await readableCvText(admin, cv);
    if (cvText.length < 80) return json({ error: "Upload a readable PDF or text CV before requesting a review." }, 400);
    const { data: allowance, error: allowanceError } = await admin.rpc("consume_document_generation_allowance", { p_user_id: user.id, p_feature: "cv_tailoring", p_limit: allowanceLimit });
    if (allowanceError) throw allowanceError;
    if (!allowance?.allowed) return json({ error: `You have used all ${allowanceLimit} CV reviews and tailored CV generations available this month. Upgrade your plan or try again next month.`, allowance }, 429);
    const jobKeywords = importantWords(`${job.title} ${job.description}`, 24);
    const cvWordSet = new Set(words(cvText));
    const matched = jobKeywords.filter((word) => cvWordSet.has(word)).slice(0, 14);
    const missing = jobKeywords.filter((word) => !cvWordSet.has(word)).slice(0, 10);
    const hasEmail = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i.test(cvText) || !!user.email;
    const hasSections = /(experience|employment|education|qualification|skills|profile|summary)/i.test(cvText);
    const keywordScore = Math.min(55, Math.round((matched.length / Math.max(8, Math.min(14, jobKeywords.length || 8))) * 55));
    const atsScore = Math.min(100, keywordScore + (hasEmail ? 10 : 0) + (hasSections ? 15 : 0) + (cvText.length >= 600 ? 10 : 5) + 10);
    const parsed = parseCvSections(cvText);
    const profileSummary = clean(profile?.experience_summary);
    const summary = profileSummary || parsed.summary || `Candidate with relevant experience and transferable capability for ${clean(job.title)}, presenting evidence aligned with the vacancy requirements.`;
    const profileSkills = Array.isArray(profile?.skills) ? profile.skills.map(clean) : clean(profile?.skills).split(/[,;|]+/).map((x) => x.trim()).filter(Boolean);
    const coreSkills = [...new Set([...matched, ...profileSkills, ...parsed.skills])].filter(Boolean).slice(0, 14).map((value) => value.replace(/^./, (c) => c.toUpperCase()));
    const name = clean(profile?.full_name) || clean(user.email?.split("@")[0]) || "Candidate";
    const contact = String(user.email || "");
    const title = clean(job.title) || "Target role";
    const sections: CvSections = {
      summary,
      skills: coreSkills.length ? coreSkills : ["Communication", "Teamwork", "Organisation"],
      experience: parsed.experience.length ? parsed.experience : [{ heading: "Relevant experience", bullets: splitItems(cvText).slice(0, 14) }],
      education: parsed.education,
      qualifications: parsed.qualifications,
      languages: parsed.languages,
    };
    const tailoredText = [
      name, contact, "", title, "", "PROFESSIONAL PROFILE", sections.summary, "", "CORE SKILLS",
      ...sections.skills.map((skill) => `• ${skill}`), "", "PROFESSIONAL EXPERIENCE",
      ...sections.experience.flatMap((block) => [block.heading, ...block.bullets.map((item) => `• ${item}`)]),
      ...(sections.education.length ? ["", "EDUCATION", ...sections.education.map((item) => `• ${item}`)] : []),
      ...(sections.qualifications.length ? ["", "QUALIFICATIONS AND TRAINING", ...sections.qualifications.map((item) => `• ${item}`)] : []),
      ...(sections.languages.length ? ["", "LANGUAGES", ...sections.languages.map((item) => `• ${item}`)] : []),
    ].join("\n").slice(0, 50000);
    const now = new Date().toISOString();
    const { data: tailored, error: insertError } = await admin.from("tailored_cv_versions").insert({
      user_id: user.id, job_id: saved.job_id, cv_id: cv.id, version_name: `${title} · ATS tailored CV`, ats_score: atsScore, tailored_text: tailoredText, created_at: now, updated_at: now,
    }).select("id,job_id,cv_id,version_name,ats_score,tailored_text,file_bucket,file_path,file_name,updated_at,created_at").single();
    if (insertError) throw insertError;
    const safeTitle = title.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 65) || "tailored-cv";
    const fileName = `${safeTitle}-ats-cv.docx`, filePath = `${user.id}/tailored-cvs/${tailored.id}-${fileName}`, pdfPath = filePath.replace(/\.docx$/i, ".pdf");
    const [docxBytes, pdfBytes] = await Promise.all([cvDocx(name, contact, title, sections), cvPdf(name, contact, title, sections)]);
    const [{ error: docxError }, { error: pdfError }] = await Promise.all([
      admin.storage.from("visapilot-documents").upload(filePath, new Blob([docxBytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }), { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", upsert: true }),
      admin.storage.from("visapilot-documents").upload(pdfPath, new Blob([pdfBytes], { type: "application/pdf" }), { contentType: "application/pdf", upsert: true }),
    ]);
    if (docxError || pdfError) throw docxError || pdfError;
    const { data: stored, error: updateError } = await admin.from("tailored_cv_versions").update({ file_bucket: "visapilot-documents", file_path: filePath, file_name: fileName, updated_at: new Date().toISOString() }).eq("id", tailored.id).select("id,job_id,cv_id,version_name,ats_score,tailored_text,file_bucket,file_path,file_name,updated_at,created_at").single();
    if (updateError) throw updateError;
    const review = { match_score: atsScore, matched_keywords: matched, missing_keywords: missing, strengths: [hasSections ? "Recognisable ATS section headings" : "Readable CV content", hasEmail ? "Candidate contact information available" : "Candidate account email available", `${matched.length} relevant vacancy terms supported by the CV`], improvements: [missing.length ? `Consider evidencing these requirements where truthful: ${missing.slice(0, 6).join(", ")}` : "Strong keyword coverage for this vacancy", "Review dates, job titles and achievements before submitting", "Keep formatting single-column and avoid graphics, tables and text boxes"] };
    await admin.from("cv_job_analyses").delete().eq("user_id", user.id).eq("job_id", saved.job_id);
    const { data: analysis } = await admin.from("cv_job_analyses").insert({ user_id: user.id, job_id: saved.job_id, cv_id: cv.id, match_score: atsScore, matched_keywords: matched, missing_keywords: missing, created_at: now }).select("id,job_id,match_score,matched_keywords,missing_keywords,created_at").maybeSingle();
    return json({ success: true, tailored_cv: stored, analysis: analysis || review, review, pdf_path: pdfPath, allowance });
  } catch (error) {
    console.error("generate-tailored-cv:", error);
    return json({ error: error instanceof Error ? error.message : "CV review and tailoring failed." }, 500);
  }
});
