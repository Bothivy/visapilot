import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import { extractText, getDocumentProxy } from "npm:unpdf@1.8.1";
import { AlignmentType, Document, Packer, Paragraph, TextRun } from "npm:docx@9.5.1";
import { PDFDocument, StandardFonts, rgb } from "npm:pdf-lib@1.17.1";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
};

const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });

const stopWords = new Set(
  "the and for with from this that your you are our their they them job role work working will would can required essential desirable experience skills skill ability team teams about using use who what when where how not but all any have has had into position candidate successful including within across through also such must should".split(" "),
);

function clean(value: unknown) {
  return String(value || "")
    .replace(/<[^>]*>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/\s+/g, " ")
    .trim();
}

function words(value: unknown) {
  return clean(value).toLowerCase().match(/[a-z][a-z0-9+#.-]{2,}/g) || [];
}

function matchedKeywords(jobText: string, cvText: string) {
  const cvWords = new Set(words(cvText));
  const frequency = new Map<string, number>();
  for (const word of words(jobText)) {
    if (stopWords.has(word) || !cvWords.has(word)) continue;
    frequency.set(word, (frequency.get(word) || 0) + 1);
  }
  return [...frequency.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8)
    .map(([word]) => word);
}

function cvEvidence(cvText: string, keywords: string[]) {
  const lines = String(cvText || "")
    .split(/(?<=[.!?])\s+|\n+/)
    .map((line) => clean(line.replace(/^[-•]\s*/, "")))
    .filter((line) =>
      line.length >= 45 &&
      line.length <= 280 &&
      !/@/.test(line) &&
      !/\+?\d[\d\s()-]{8,}/.test(line)
    );
  const scored = lines.map((line) => {
    const lower = line.toLowerCase();
    return { line, score: keywords.filter((word) => lower.includes(word)).length };
  });
  return scored
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, 2)
    .map((item) => item.line);
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

async function coverLetterDocx(text: string) {
  const blocks = String(text).split(/\n{2,}/).map((value) => value.trim()).filter(Boolean);
  const document = new Document({
    sections: [{
      properties: { page: { margin: { top: 900, right: 900, bottom: 900, left: 900 } } },
      children: blocks.map((block, index) => new Paragraph({
        alignment: index === 0 || index === blocks.length - 1 ? AlignmentType.LEFT : AlignmentType.JUSTIFIED,
        spacing: { after: 220, line: 300 },
        children: [new TextRun({ text: block, size: 22, font: "Arial" })],
      })),
    }],
  });
  return new Uint8Array(await Packer.toBuffer(document));
}

function wrappedLines(text: string, font: any, size: number, width: number) {
  const result: string[] = [];
  for (const paragraph of String(text).split(/\n/)) {
    if (!paragraph.trim()) { result.push(""); continue; }
    const words = paragraph.trim().split(/\s+/);
    let line = "";
    for (const word of words) {
      const candidate = line ? `${line} ${word}` : word;
      if (font.widthOfTextAtSize(candidate, size) <= width) line = candidate;
      else { if (line) result.push(line); line = word; }
    }
    if (line) result.push(line);
  }
  return result;
}

async function coverLetterPdf(text: string) {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const size = 11;
  const margin = 58;
  const lineHeight = 17;
  let page = pdf.addPage([595.28, 841.89]);
  let y = page.getHeight() - margin;
  for (const line of wrappedLines(text, font, size, page.getWidth() - margin * 2)) {
    if (y < margin) { page = pdf.addPage([595.28, 841.89]); y = page.getHeight() - margin; }
    if (line) page.drawText(line, { x: margin, y, size, font, color: rgb(0.09, 0.08, 0.16) });
    y -= line ? lineHeight : lineHeight * 0.75;
  }
  return new Uint8Array(await pdf.save());
}

async function readableCvText(admin: any, cv: any) {
  const existing = clean(cv?.cv_text);
  if (existing.length >= 80) return { text: existing, source: "saved_cv_text" };
  if (!cv?.storage_bucket || !cv?.storage_path) return { text: "", source: "none" };

  const { data: file, error } = await admin.storage
    .from(cv.storage_bucket)
    .download(cv.storage_path);
  if (error || !file) return { text: "", source: "none" };

  const filename = String(cv?.original_filename || cv?.storage_path || "").toLowerCase();
  let extracted = "";
  if (filename.endsWith(".pdf")) {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const pdf = await getDocumentProxy(bytes);
    const result = await extractText(pdf, { mergePages: true });
    extracted = clean(result?.text);
  } else if (filename.endsWith(".txt")) {
    extracted = clean(await file.text());
  } else if (filename.endsWith(".rtf")) {
    extracted = clean(
      (await file.text())
        .replace(/\\'[0-9a-f]{2}/gi, " ")
        .replace(/\\[a-z]+-?\d* ?/gi, " ")
        .replace(/[{}]/g, " "),
    );
  }

  if (extracted.length >= 80) {
    await admin.from("candidate_cvs").update({
      cv_text: extracted.slice(0, 50000),
      updated_at: new Date().toISOString(),
    }).eq("id", cv.id);
    return { text: extracted, source: "uploaded_cv_file" };
  }
  return { text: "", source: "none" };
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  if (req.method !== "POST") return json({ error: "Method not allowed." }, 405);

  try {
    const supabaseUrl = Deno.env.get("SUPABASE_URL");
    const anonKey = Deno.env.get("SUPABASE_ANON_KEY");
    const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
    if (!supabaseUrl || !anonKey || !serviceKey) {
      throw new Error("Missing required environment variable.");
    }

    const authorization = req.headers.get("Authorization") || "";
    if (!authorization.startsWith("Bearer ")) return json({ error: "Unauthorized" }, 401);
    const candidate = createClient(supabaseUrl, anonKey, {
      global: { headers: { Authorization: authorization } },
      auth: { persistSession: false, autoRefreshToken: false },
    });
    const { data: { user }, error: userError } = await candidate.auth.getUser();
    if (userError || !user) return json({ error: "Unauthorized" }, 401);

    const admin = createClient(supabaseUrl, serviceKey, {
      auth: { persistSession: false, autoRefreshToken: false },
    });
    const body = await req.json().catch(() => ({}));
    const existingLetterId = String(body?.cover_letter_id || "");
    if (body?.action === "create_files" && existingLetterId) {
      const { data: existingLetter, error: existingError } = await admin
        .from("cover_letter_versions")
        .select("id,user_id,job_id,version_name,cover_letter_text,file_bucket,file_path,file_name,updated_at,created_at")
        .eq("id", existingLetterId)
        .eq("user_id", user.id)
        .maybeSingle();
      if (existingError) throw existingError;
      if (!existingLetter) return json({ error: "Cover letter not found." }, 404);
      if (!clean(existingLetter.cover_letter_text)) return json({ error: "This cover letter has no readable text to convert." }, 400);
      const safeName = clean(existingLetter.version_name || "cover-letter").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 70) || "cover-letter";
      const fileName = `${safeName}.docx`;
      const filePath = `${user.id}/cover-letters/${existingLetter.id}-${fileName}`;
      const pdfPath = filePath.replace(/\.docx$/i, ".pdf");
      const [docxBytes, pdfBytes] = await Promise.all([
        coverLetterDocx(existingLetter.cover_letter_text),
        coverLetterPdf(existingLetter.cover_letter_text),
      ]);
      const [{ error: docxError }, { error: pdfError }] = await Promise.all([
        admin.storage.from("visapilot-documents").upload(filePath, new Blob([docxBytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }), { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", upsert: true }),
        admin.storage.from("visapilot-documents").upload(pdfPath, new Blob([pdfBytes], { type: "application/pdf" }), { contentType: "application/pdf", upsert: true }),
      ]);
      if (docxError || pdfError) throw docxError || pdfError;
      const { data: updatedLetter, error: updateError } = await admin
        .from("cover_letter_versions")
        .update({ file_bucket: "visapilot-documents", file_path: filePath, file_name: fileName, updated_at: new Date().toISOString() })
        .eq("id", existingLetter.id)
        .select("id,job_id,version_name,cover_letter_text,file_bucket,file_path,file_name,updated_at,created_at")
        .single();
      if (updateError) throw updateError;
      return json({ success: true, cover_letter: updatedLetter, converted_existing: true });
    }
    const savedJobId = String(body?.saved_job_id || "");
    if (!savedJobId) return json({ error: "A saved application is required." }, 400);

    const { data: billing, error: billingError } = await admin
      .from("billing_subscriptions")
      .select("status,current_period_end,plan_key,stripe_price_id")
      .eq("user_id", user.id)
      .maybeSingle();
    if (billingError) throw billingError;
    const planKey = String(billing?.plan_key || planFromPrice(String(billing?.stripe_price_id || "")));
    const end = billing?.current_period_end ? new Date(billing.current_period_end).getTime() : null;
    const activePlan = ["active", "trialing"].includes(String(billing?.status || "")) &&
      (!end || Number.isNaN(end) || end > Date.now());
    const allowanceLimit = activePlan && /^pro_/.test(planKey)
      ? -1
      : activePlan && /^copilot_/.test(planKey)
      ? 25
      : 5;

    const { data: saved, error: savedError } = await admin
      .from("saved_jobs")
      .select("id,job_id,user_id")
      .eq("id", savedJobId)
      .eq("user_id", user.id)
      .maybeSingle();
    if (savedError) throw savedError;
    if (!saved) return json({ error: "Saved application not found." }, 404);

    const [{ data: job, error: jobError }, { data: cv, error: cvError }, { data: profile, error: profileError }, { data: tailored, error: tailoredError }] =
      await Promise.all([
        admin.from("jobs").select("id,title,location,description,employer_id,employers(name)").eq("id", saved.job_id).maybeSingle(),
        admin.from("candidate_cvs").select("id,cv_name,original_filename,cv_text,storage_bucket,storage_path,updated_at").eq("user_id", user.id).order("updated_at", { ascending: false }).limit(1).maybeSingle(),
        admin.from("candidate_profiles").select("full_name,skills,experience_summary,target_roles").eq("user_id", user.id).maybeSingle(),
        admin.from("tailored_cv_versions").select("tailored_text,updated_at").eq("user_id", user.id).order("updated_at", { ascending: false }).limit(1).maybeSingle(),
      ]);
    if (jobError) throw jobError;
    if (cvError) throw cvError;
    if (profileError) throw profileError;
    if (tailoredError) throw tailoredError;
    if (!job) return json({ error: "Job details could not be found." }, 404);

    let cvSource = "none";
    let cvText = "";
    if (cv) {
      try {
        const readable = await readableCvText(admin, cv);
        cvText = readable.text;
        cvSource = readable.source;
      } catch (extractError) {
        console.warn("generate-cover-letter: CV extraction fallback", extractError);
      }
    }
    if (cvText.length < 80 && clean(tailored?.tailored_text).length >= 80) {
      cvText = clean(tailored.tailored_text);
      cvSource = "tailored_cv";
    }
    if (cvText.length < 80) {
      cvText = clean([profile?.experience_summary, profile?.skills, profile?.target_roles].filter(Boolean).join(". "));
      cvSource = "candidate_profile";
    }
    if (cvText.length < 30) {
      return json({ error: "Add experience and skills to My Profile, or upload a readable PDF or text CV." }, 400);
    }

    const { data: allowance, error: allowanceError } = await admin.rpc(
      "consume_document_generation_allowance",
      { p_user_id: user.id, p_feature: "cover_letter", p_limit: allowanceLimit },
    );
    if (allowanceError) throw allowanceError;
    if (!allowance?.allowed) {
      return json({ error: `You have used all ${allowanceLimit} cover-letter generations available this month. Upgrade your plan or try again next month.`, allowance }, 429);
    }

    const title = clean(job.title) || "the advertised role";
    const employer = clean(job?.employers?.name) || "your organisation";
    const candidateName = clean(profile?.full_name) || clean(user.email?.split("@")[0]) || "Candidate";
    const jobText = `${title} ${clean(job.description)}`;
    const keywords = matchedKeywords(jobText, cvText);
    const evidence = cvEvidence(cvText, keywords);
    const skillText = keywords.length
      ? keywords.slice(0, 6).join(", ")
      : clean(profile?.skills).split(/[,;]+/).map((item) => item.trim()).filter(Boolean).slice(0, 6).join(", ");

    const paragraphs = [
      "Dear Hiring Manager,",
      `I am writing to apply for the ${title} position at ${employer}. My experience and CV demonstrate a strong interest in this opportunity and relevant capability for the responsibilities described.`,
    ];
    const summary = clean(profile?.experience_summary);
    if (summary) paragraphs.push(summary);
    if (skillText) {
      paragraphs.push(`My background aligns particularly well with the role’s focus on ${skillText}. I would bring this relevant experience to the position while continuing to learn the organisation’s systems and ways of working.`);
    }
    if (evidence.length) {
      paragraphs.push(`Relevant evidence from my CV includes ${evidence.map((item) => item.replace(/[.!?]+$/, "")).join(". In addition, ")}.`);
    }
    paragraphs.push(
      `I would welcome the opportunity to discuss how my experience could contribute to ${employer}. Thank you for considering my application.`,
      `Kind regards,\n${candidateName}`,
    );
    const coverLetter = paragraphs.join("\n\n").slice(0, 6000);
    const now = new Date().toISOString();
    const { data: inserted, error: insertError } = await admin
      .from("cover_letter_versions")
      .insert({
        user_id: user.id,
        job_id: saved.job_id,
        version_name: `${title} · tailored cover letter`,
        cover_letter_text: coverLetter,
        created_at: now,
        updated_at: now,
      })
      .select("id,job_id,version_name,cover_letter_text,file_bucket,file_path,file_name,updated_at,created_at")
      .single();
    if (insertError) throw insertError;

    const safeTitle = title
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 70) || "cover-letter";
    const fileName = `${safeTitle}-cover-letter.docx`;
    const filePath = `${user.id}/cover-letters/${inserted.id}-${fileName}`;
    const pdfPath = filePath.replace(/\.docx$/i, ".pdf");
    const [docxBytes, pdfBytes] = await Promise.all([
      coverLetterDocx(coverLetter),
      coverLetterPdf(coverLetter),
    ]);
    const [{ error: uploadError }, { error: pdfUploadError }] = await Promise.all([
      admin.storage.from("visapilot-documents").upload(
        filePath,
        new Blob([docxBytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }),
        { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", upsert: true },
      ),
      admin.storage.from("visapilot-documents").upload(
        pdfPath,
        new Blob([pdfBytes], { type: "application/pdf" }),
        { contentType: "application/pdf", upsert: true },
      ),
    ]);

    let storedLetter = inserted;
    let storageWarning: string | null = null;
    if (uploadError || pdfUploadError) {
      console.error("generate-cover-letter: private storage upload", uploadError || pdfUploadError);
      storageWarning = "The letter was generated, but its private file could not be created.";
    } else {
      const { data: updated, error: updateError } = await admin
        .from("cover_letter_versions")
        .update({
          file_bucket: "visapilot-documents",
          file_path: filePath,
          file_name: fileName,
          updated_at: new Date().toISOString(),
        })
        .eq("id", inserted.id)
        .select("id,job_id,version_name,cover_letter_text,file_bucket,file_path,file_name,updated_at,created_at")
        .single();
      if (updateError) throw updateError;
      storedLetter = updated;
    }

    return json({
      success: true,
      cover_letter: storedLetter,
      matched_keywords: keywords,
      evidence_source: cvSource,
      storage_warning: storageWarning,
      allowance,
    });
  } catch (error) {
    console.error("generate-cover-letter:", error);
    return json(
      { error: error instanceof Error ? error.message : "Cover letter generation failed." },
      500,
    );
  }
});
