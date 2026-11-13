// Stand-in for python/bridge.py that speaks the same line-delimited JSON protocol.
// Lets the client tests provoke hangs, crashes and protocol noise on demand.
import { randomUUID } from "node:crypto";
import { createInterface } from "node:readline";

function reply(id, result) {
  process.stdout.write(`${JSON.stringify({ id, result })}\n`);
}

createInterface({ input: process.stdin }).on("line", (line) => {
  const { id, method, params } = JSON.parse(line);

  switch (method) {
    case "pptx_get_engine_info":
      reply(id, { engine: "fake", pid: process.pid });
      break;
    case "pptx_create_presentation":
      reply(id, { presentation_id: randomUUID() });
      break;
    case "hang":
      break;
    case "crash":
      process.stderr.write("fake bridge crashing on request\n");
      process.exit(3);
      break;
    case "noise":
      process.stdout.write("this line is not JSON\n");
      reply(id, { ok: true });
      break;
    case "fail":
      process.stdout.write(`${JSON.stringify({ id, error: { code: "not_found", message: "nope" } })}\n`);
      break;
    case "__shutdown__":
      reply(id, { ok: true });
      process.exit(0);
      break;
    default:
      reply(id, { method, params });
  }
});
