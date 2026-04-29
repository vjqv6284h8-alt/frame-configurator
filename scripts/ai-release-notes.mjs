import { readFileSync } from "node:fs";

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

async function generateReleaseNotes({ apiKey, model, prompt }) {
  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model,
      input: prompt,
      max_output_tokens: 1200,
    }),
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`OpenAI API ${response.status}: ${body}`);
  }

  const data = await response.json();
  return (data.output_text || "").trim();
}

function buildPrompt({ repository, tag, previousTag, commitDigest }) {
  return [
    "Ты release manager для React + TypeScript + Express проекта.",
    "Сделай понятные release notes на русском языке.",
    "Верни markdown со строго такой структурой:",
    "## Что вошло",
    "- 3-8 пунктов по изменениям человеческим языком",
    "## Важные проверки",
    "- список проверок после деплоя",
    "## Технические детали",
    "- кратко: совместимость, риски, что мониторить",
    "",
    "Требования:",
    "- Не выдумывай изменения, опирайся только на входные данные.",
    "- Если данных мало, явно напиши, что информации недостаточно.",
    "- Избегай общих фраз и воды.",
    "",
    `Repository: ${repository}`,
    `Current tag: ${tag}`,
    `Previous tag: ${previousTag || "(first release or unknown)"}`,
    "",
    "Commits and stats:",
    commitDigest || "(no commits data)",
  ].join("\n");
}

async function main() {
  const githubToken = required("GITHUB_TOKEN");
  const openaiApiKey = required("OPENAI_API_KEY");
  const model = process.env.OPENAI_MODEL || "gpt-4.1-mini";
  const eventPath = required("GITHUB_EVENT_PATH");
  const event = JSON.parse(readFileSync(eventPath, "utf8"));
  const release = event.release;

  if (!release) {
    console.log("No release payload, skipping.");
    return;
  }

  const commitDigest = readFileSync(0, "utf8");
  const prompt = buildPrompt({
    repository: event.repository?.full_name || "unknown",
    tag: release.tag_name,
    previousTag: process.env.PREVIOUS_TAG || "",
    commitDigest,
  });

  const notes = await generateReleaseNotes({
    apiKey: openaiApiKey,
    model,
    prompt,
  });

  const body = `${notes || "Модель не вернула текст release notes."}\n\n_Model: ${model}_`;

  await ghRequest(release.url, githubToken, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ body }),
  });

  console.log(`Updated release notes for ${release.tag_name}.`);
}

main().catch((error) => {
  console.error(error.message || error);
  process.exit(1);
});
