import { z } from "zod";

const uuid = z.string().uuid();
/** Runtime schema for `sommelier.config.ts`. */
export const sommelierConfigSchema = z.object({
  app: z.object({
    id: z.union([z.literal("generate"), uuid]),
    name: z.string().min(1).max(125),
    description: z.string().min(1).max(250),
    version: z
      .string()
      .regex(/^\d+\.\d+\.\d+\.\d+$|^\d+\.\d+\.\d+$/, "Use a dotted numeric version.")
      .refine((value) => {
        const parts = value.split(".").map(Number);
        return (parts[0] ?? 99_999) <= 99_998 && parts.every((part) => part <= 99_999);
      }, "Use components up to 99999 and a major version up to 99998."),
    providerName: z.string().min(1).max(125).default("Sommelier contributors"),
    locale: z
      .string()
      .regex(/^[a-z]{2}-[A-Z]{2}$/)
      .default("en-US"),
  }),
  office: z.object({
    hosts: z.tuple([z.literal("Workbook")]),
    permissions: z.literal("ReadWriteDocument"),
    requirements: z.object({ ExcelApi: z.string().regex(/^\d+\.\d+$/) }),
  }),
  taskpane: z.object({
    developmentUrl: z.string().url(),
    productionUrl: z.string().url(),
    developmentCommand: z.string().min(1).optional(),
    entryPath: z
      .string()
      .regex(/^\/(?!\/)[a-zA-Z0-9/_.-]*$/)
      .optional(),
    supportPath: z
      .string()
      .regex(/^\/(?!\/)[a-zA-Z0-9/_.-]*$/)
      .optional(),
  }),
  commands: z.object({
    label: z.string().min(1).max(80),
    groupLabel: z.string().min(1).max(80).default("Sommelier"),
    icon: z.string().min(1),
    icons: z
      .object({
        16: z.string().min(1),
        32: z.string().min(1),
        64: z.string().min(1),
        80: z.string().min(1),
      })
      .optional(),
  }),
});

/** Typed Sommelier project configuration. */
export type SommelierConfig = z.input<typeof sommelierConfigSchema>;
/** Validated configuration with defaults applied. */
export type ResolvedSommelierConfig = z.output<typeof sommelierConfigSchema>;

/** Preserve config inference while runtime loading performs validation. */
export function defineConfig(config: SommelierConfig): SommelierConfig {
  return config;
}

/** Published JSON Schema for editor integration. */
export const sommelierConfigJsonSchema = z.toJSONSchema(sommelierConfigSchema, {
  target: "draft-7",
});
