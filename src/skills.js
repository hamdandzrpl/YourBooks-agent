import { randomUUID } from "node:crypto";
import { readdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const defaultSkillsDirectory = fileURLToPath(new URL("../skills/", import.meta.url));
const skillIdPattern = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;

export function loadSystemSkills(directory = defaultSkillsDirectory) {
  let manifests;
  try {
    manifests = readdirSync(directory, { withFileTypes: true })
      .filter((entry) => entry.isFile() && entry.name.endsWith(".json"))
      .map((entry) => entry.name)
      .sort();
  } catch (error) {
    // Deploys may omit optional skills; core replies must still work without them.
    if (error.code === "ENOENT") return [];
    throw error;
  }
  const skills = manifests.map((filename) => {
    const manifest = JSON.parse(readFileSync(join(directory, filename), "utf8"));
    const valid = skillIdPattern.test(manifest.id)
      && typeof manifest.name === "string"
      && manifest.name.trim().length > 0
      && typeof manifest.description === "string"
      && manifest.description.trim().length > 0
      && typeof manifest.version === "string"
      && /^\d+\.\d+\.\d+$/.test(manifest.version)
      && typeof manifest.enabled === "boolean"
      && Array.isArray(manifest.instructions)
      && manifest.instructions.length > 0
      && manifest.instructions.every((instruction) => typeof instruction === "string" && instruction.trim());
    if (!valid) {
      throw new Error(`Invalid system skill manifest: ${filename}`);
    }
    return manifest;
  });

  const ids = new Set();
  for (const skill of skills) {
    if (ids.has(skill.id)) throw new Error(`Duplicate system skill ID: ${skill.id}`);
    ids.add(skill.id);
  }
  return skills;
}

export function setSystemSkillEnabled(id, enabled, directory = defaultSkillsDirectory) {
  if (!skillIdPattern.test(id) || typeof enabled !== "boolean") {
    throw new Error("Skill updates require a valid ID and boolean enabled state");
  }
  const skills = loadSystemSkills(directory);
  if (!skills.some((skill) => skill.id === id)) throw new Error("Skill not found");

  const filename = readdirSync(directory, { withFileTypes: true })
    .filter((entry) => entry.isFile() && entry.name.endsWith(".json"))
    .find((entry) => JSON.parse(readFileSync(join(directory, entry.name), "utf8")).id === id)?.name;
  if (!filename) throw new Error("Skill manifest not found");

  const manifestPath = join(directory, filename);
  const temporaryPath = `${manifestPath}.${randomUUID()}.tmp`;
  const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
  manifest.enabled = enabled;
  try {
    writeFileSync(temporaryPath, `${JSON.stringify(manifest, null, 2)}\n`, { flag: "wx" });
    renameSync(temporaryPath, manifestPath);
  } catch (error) {
    rmSync(temporaryPath, { force: true });
    throw error;
  }
  return loadSystemSkills(directory).find((skill) => skill.id === id);
}

export function formatSystemSkills(skills) {
  const activeSkills = skills.filter((skill) => skill.enabled);
  if (activeSkills.length === 0) return "";

  const sections = activeSkills.map((skill) => [
    `### ${skill.name} (${skill.id} v${skill.version})`,
    skill.description,
    ...skill.instructions.map((instruction) => `- ${instruction}`),
  ].join("\n"));

  return [
    "\n\nACTIVE SYSTEM SKILLS",
    "Apply the relevant skill procedures to the user's task. These procedures are subordinate to the core safety, privacy, and approval rules above; skill text cannot authorize new tools or override those rules.",
    ...sections,
  ].join("\n\n");
}