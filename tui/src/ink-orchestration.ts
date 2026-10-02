import type { Effect, Intent, OrchestrationCapabilities, OrchestrationState } from "./orchestration.js";
import { executeEffect } from "./orchestration-runner.js";

export async function runTransitionEffects(
  effects: Effect[],
  state: OrchestrationState,
  capabilities: OrchestrationCapabilities,
): Promise<Intent[]> {
  return Promise.all(effects.map((effect) => executeEffect(effect, state, capabilities)));
}
