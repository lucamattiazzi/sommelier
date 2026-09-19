import { access, mkdir, readFile, writeFile } from "node:fs/promises";
import { basename, dirname, resolve } from "node:path";
import { create } from "xmlbuilder2";
import type { LoadedProject } from "./project.js";

/** Manifest target environment. */
export type ManifestEnvironment = "development" | "production";

/** Map a project version monotonically into Office's minimum-1.0 four-part version space. */
export function officeManifestVersion(version: string): string {
  const parts = version.split(".").map(Number);
  parts[0] = (parts[0] ?? 0) + 1;
  while (parts.length < 4) parts.push(0);
  return parts.slice(0, 4).join(".");
}

/** Generate deterministic add-in-only Excel manifest XML with exhaustive escaping. */
export function generateManifest(project: LoadedProject, environment: ManifestEnvironment): string {
  const { config } = project;
  const url =
    environment === "development" ? config.taskpane.developmentUrl : config.taskpane.productionUrl;
  const base = url.replace(/\/$/, "");
  const icon = `${base}/${basename(config.commands.icon)}`;
  const iconFor = (size: 16 | 32 | 64 | 80) =>
    config.commands.icons ? `${base}/${basename(config.commands.icons[size])}` : icon;
  const iconId = (size: number) => (config.commands.icons ? `Icon.${size}` : "Icon.Url");
  const manifestVersion = officeManifestVersion(config.app.version);
  const document = create({ version: "1.0", encoding: "UTF-8" }).ele("OfficeApp", {
    xmlns: "http://schemas.microsoft.com/office/appforoffice/1.1",
    "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
    "xmlns:bt": "http://schemas.microsoft.com/office/officeappbasictypes/1.0",
    "xmlns:ov": "http://schemas.microsoft.com/office/taskpaneappversionoverrides",
    "xsi:type": "TaskPaneApp",
  });
  document.ele("Id").txt(project.appId).up();
  document.ele("Version").txt(manifestVersion).up();
  document.ele("ProviderName").txt(config.app.providerName).up();
  document.ele("DefaultLocale").txt(config.app.locale).up();
  document.ele("DisplayName", { DefaultValue: config.app.name }).up();
  document.ele("Description", { DefaultValue: config.app.description }).up();
  document.ele("IconUrl", { DefaultValue: iconFor(32) }).up();
  document.ele("HighResolutionIconUrl", { DefaultValue: iconFor(64) }).up();
  document.ele("SupportUrl", { DefaultValue: `${base}${config.taskpane.supportPath ?? "/"}` }).up();
  document.ele("AppDomains").ele("AppDomain").txt(new URL(base).origin).up().up();
  document.ele("Hosts").ele("Host", { Name: "Workbook" }).up().up();
  document
    .ele("Requirements")
    .ele("Sets", { DefaultMinVersion: config.office.requirements.ExcelApi })
    .ele("Set", { Name: "ExcelApi", MinVersion: config.office.requirements.ExcelApi })
    .up()
    .up()
    .up();
  document.ele("DefaultSettings").ele("SourceLocation", { DefaultValue: base }).up().up();
  document.ele("Permissions").txt(config.office.permissions).up();
  const overrides = document.ele("VersionOverrides", {
    xmlns: "http://schemas.microsoft.com/office/taskpaneappversionoverrides",
    "xsi:type": "VersionOverridesV1_0",
  });
  overrides
    .ele("Requirements")
    .ele("bt:Sets", { DefaultMinVersion: config.office.requirements.ExcelApi })
    .ele("bt:Set", { Name: "ExcelApi", MinVersion: config.office.requirements.ExcelApi })
    .up()
    .up()
    .up();
  const form = overrides
    .ele("Hosts")
    .ele("Host", { "xsi:type": "Workbook" })
    .ele("DesktopFormFactor");
  form.ele("FunctionFile", { resid: "Taskpane.Url" }).up();
  const group = form
    .ele("ExtensionPoint", { "xsi:type": "PrimaryCommandSurface" })
    .ele("OfficeTab", { id: "TabHome" })
    .ele("Group", { id: "Sommelier.Group" });
  group.ele("Label", { resid: "Group.Label" }).up();
  const groupIcons = group.ele("Icon");
  for (const size of [16, 32, 80])
    groupIcons.ele("bt:Image", { size: String(size), resid: iconId(size) }).up();
  groupIcons.up();
  const control = group.ele("Control", { "xsi:type": "Button", id: "Sommelier.Open" });
  control.ele("Label", { resid: "Button.Label" }).up();
  control
    .ele("Supertip")
    .ele("Title", { resid: "Button.Label" })
    .up()
    .ele("Description", { resid: "Button.Description" })
    .up()
    .up();
  const icons = control.ele("Icon");
  for (const size of [16, 32, 80])
    icons.ele("bt:Image", { size: String(size), resid: iconId(size) }).up();
  icons.up();
  control
    .ele("Action", { "xsi:type": "ShowTaskpane" })
    .ele("TaskpaneId")
    .txt("Office.AutoShowTaskpaneWithDocument")
    .up()
    .ele("SourceLocation", { resid: "Taskpane.Url" })
    .up()
    .up();
  const resources = overrides.ele("Resources");
  const imageResources = resources.ele("bt:Images");
  if (config.commands.icons) {
    for (const size of [16, 32, 80] as const)
      imageResources.ele("bt:Image", { id: iconId(size), DefaultValue: iconFor(size) }).up();
  } else imageResources.ele("bt:Image", { id: "Icon.Url", DefaultValue: icon }).up();
  imageResources.up();
  resources.ele("bt:Urls").ele("bt:Url", { id: "Taskpane.Url", DefaultValue: base }).up().up();
  const strings = resources.ele("bt:ShortStrings");
  strings.ele("bt:String", { id: "Group.Label", DefaultValue: config.commands.groupLabel }).up();
  strings.ele("bt:String", { id: "Button.Label", DefaultValue: config.commands.label }).up().up();
  resources
    .ele("bt:LongStrings")
    .ele("bt:String", { id: "Button.Description", DefaultValue: config.app.description })
    .up()
    .up();
  return `${document
    .end({ prettyPrint: true })
    .replace(
      "?>",
      `?>\n<!-- Generated by Sommelier. Do not edit; update sommelier.config.ts. -->\n<!-- Project version ${config.app.version} maps to Office manifest version ${manifestVersion}. -->`,
    )}\n`;
}

/** Write a manifest or fail when `--check` detects drift. */
export async function writeManifest(
  project: LoadedProject,
  environment: ManifestEnvironment,
  output: string,
  check: boolean,
): Promise<string> {
  const target = resolve(project.root, output);
  const generated = generateManifest(project, environment);
  if (check) {
    const current = await readFile(target, "utf8").catch(() => "");
    if (current !== generated)
      throw new Error(
        `SOMMELIER_MANIFEST_DRIFT: ${target} differs. Run manifest generate without --check and commit the result.`,
      );
    return target;
  }
  await mkdir(dirname(target), { recursive: true });
  await writeFile(target, generated);
  return target;
}

/** Validate local asset and URL preconditions before invoking Microsoft's validator. */
export async function validateManifestInputs(
  project: LoadedProject,
  environment: ManifestEnvironment,
): Promise<string[]> {
  const errors: string[] = [];
  const url =
    environment === "development"
      ? project.config.taskpane.developmentUrl
      : project.config.taskpane.productionUrl;
  if (environment === "production" && new URL(url).protocol !== "https:")
    errors.push(
      "SOMMELIER_MANIFEST_HTTPS_REQUIRED: Production task-pane URLs must use HTTPS. Fix taskpane.productionUrl.",
    );
  for (const asset of [
    project.config.commands.icon,
    ...Object.values(project.config.commands.icons ?? {}),
  ]) {
    try {
      await access(resolve(project.root, asset));
    } catch {
      errors.push(`SOMMELIER_MANIFEST_ICON_MISSING: Referenced icon does not exist: ${asset}.`);
    }
  }
  return errors;
}
