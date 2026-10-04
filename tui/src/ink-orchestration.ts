import type { Effect, Intent, OrchestrationCapabilities, OrchestrationState } from "./orchestration.js";
import { executeEffect } from "./orchestration-runner.js";
import { update } from "./orchestration.js";

export async function runTransitionEffects(
  effects: Effect[],
  state: OrchestrationState,
  capabilities: OrchestrationCapabilities,
): Promise<Intent[]> {
  return Promise.all(effects.map((effect) => executeEffect(effect, state, capabilities)));
}

// State is advanced synchronously, so consecutive input events see pending work
// immediately. Effects run once here, never in a React state updater.
export function createOrchestrationController(
  initialState: OrchestrationState,
  capabilities: OrchestrationCapabilities,
  onState: (state: OrchestrationState) => void,
) {
  let state = initialState;
  let disposed = false;
  async function dispatch(intent: Intent): Promise<void> {
    if (disposed) return;
    const transition = update(state, intent);
    state = transition.state;
    onState(state);
    const results = await runTransitionEffects(transition.effects, state, capabilities);
    for (const result of results) {
      if (disposed) return;
      await dispatch(result);
    }
  }
  return {
    dispatch,
    getState: () => state,
    dispose: () => { disposed = true; },
  };
}
