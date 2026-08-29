import { describe, expect, it } from "vitest";
import { verifyPaddleSignature } from "./signature";

describe("Paddle webhook verification", () => {
  it("rejects missing and forged signatures", async () => {
    expect(await verifyPaddleSignature("{}", null, "secret")).toBe(false);
    expect(await verifyPaddleSignature("{}", "ts=1;h1=forged", "secret")).toBe(false);
  });
});
