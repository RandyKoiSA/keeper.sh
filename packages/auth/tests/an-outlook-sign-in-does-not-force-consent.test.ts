import { PGlite } from "@electric-sql/pglite";
import type { BunSQLDatabase } from "drizzle-orm/bun-sql";
import { drizzle } from "drizzle-orm/pglite";
import { describe, expect, it } from "vitest";
import { createAuth } from "../src/index";

const BASE_URL = "http://localhost:3000";
const SECRET = "test-secret-for-outlook-sign-in-consent-prompt";

const buildSignInUrl = async (providerId: string): Promise<URL> => {
  const { auth } = createAuth({
    baseUrl: BASE_URL,
    commercialMode: true,
    database: drizzle(new PGlite()) as unknown as BunSQLDatabase,
    googleClientId: "test-google-client-id",
    googleClientSecret: "test-google-client-secret",
    microsoftClientId: "test-microsoft-client-id",
    microsoftClientSecret: "test-microsoft-client-secret",
    secret: SECRET,
  });

  const context = await auth.$context;
  const provider = context.socialProviders.find(
    (candidate) => candidate.id === providerId,
  );

  if (!provider) {
    throw new TypeError(`Social provider ${providerId} is not configured`);
  }

  return provider.createAuthorizationURL({
    codeVerifier: "test-code-verifier-test-code-verifier-test-code",
    redirectURI: `${BASE_URL}/api/auth/callback/${providerId}`,
    state: "test-state",
  });
};

describe("an outlook sign-in", () => {
  it("offers account selection without forcing an already-granted consent screen", async () => {
    const { searchParams } = await buildSignInUrl("microsoft");

    expect(searchParams.get("prompt")).toBe("select_account");
  });
});

describe("a google sign-in", () => {
  it("keeps forcing consent so a refresh token is always issued", async () => {
    const { searchParams } = await buildSignInUrl("google");

    expect(searchParams.get("prompt")).toBe("consent");
  });
});
