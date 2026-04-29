import { readFileSync } from "node:fs";

const marker = "<!-- ai-pr-review -->";
const maxDiffChars = 14000;

function required(name) {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing required env: ${name}`);
  }
  return value;
}

async function ghRequest(url, token, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${token}`,
      "X-GitHub-Api-Version": "2022-11-28",
      ...(options.headers || {}),
    },
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`GitHub API ${response.status}: ${body}`);
  }

  if (response.status === 204) return null;
  return response.json();
}

async function openaiReview({ apiKey, model, prompt }) {
  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model,
      input: prompt,
      max_output_tokens: 900,
    }),
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`OpenAI API ${response.status}: ${body}`);
  }

  const data = await response.json();
  return (data.output_text || "").trim();
}

function buildPrompt({ title, body, changedFiles, diff }) {
  return [
    "Ты senior reviewer для React + TypeScript + Express проекта.",
    "Дай короткий и практичный PR review на русском.",
    "Верни строго в формате:",
    "1) Риски/баги",
    "2) Что обязательно протестировать",
    "3) Рекомендации по улучшению (коротко)",
    "Если критичных проблем нет, так и напиши.",
    "",
    `PR title: ${title || "(no title)"}`,
    `PR body: ${body || "(no body)"}`,
    `Changed files: ${changedFiles}`,
    "",
    "Diff:",
    diff,
  ].join("\n");
}

function truncateDiff(diffText) {
  if (diffText.length <= maxDiffChars) return diffText;
  return `${diffText.slice(0, maxDiffChars)}\n\n[diff truncated]`;
}

async function main() {
  const githubToken = required("GITHUB_TOKEN");
  const openaiApiKey = required("OPENAI_API_KEY");
  const model = process.env.OPENAI_MODEL || "gpt-4.1-mini";
  const eventPath = required("GITHUB_EVENT_PATH");
  const event = JSON.parse(readFileSync(eventPath, "utf8"));
  const pr = event.pull_request;

  if (!pr) {
    console.log("No pull_request payload, skipping.");
    return;
  }

  const diffText = readFileSync(0, "utf8");
  const safeDiff = truncateDiff(diffText);
  const prompt = buildPrompt({
    title: pr.title,
    body: pr.body,
    changedFiles: pr.changed_files,
    diff: safeDiff,
  });

  const review = await openaiReview({
    apiKey: openaiApiKey,
    model,
    prompt,
  });

  const owner = event.repository.owner.login;
  const repo = event.repository.name;
  const issueNumber = pr.number;
  const commentsUrl = `https://api.github.com/repos/${owner}/${repo}/issues/${issueNumber}/comments`;
  const comments = await ghRequest(commentsUrl, githubToken);

  const body = [
    marker,
    "## AI Review",
    review || "Модель не вернула текст. Проверьте параметры запроса.",
    "",
    `_Model: ${model}_`,
  ].join("\n");

  const existing = comments.find((c) => typeof c.body === "string" && c.body.includes(marker));
  if (existing) {
    await ghRequest(existing.url, githubToken, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ body }),
    });
    console.log("Updated existing AI review comment.");
    return;
  }

  await ghRequest(commentsUrl, githubToken, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ body }),
  });
  console.log("Created AI review comment.");
}

main().catch((error) => {
  console.error(error.message || error);
  process.exit(1);
});
