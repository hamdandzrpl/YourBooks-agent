export class ProgressTimeline {
  constructor({ to, sendText, editText, intervalMs = 1800, editIntervalMs = 1800, setIntervalImpl = setInterval, clearIntervalImpl = clearInterval }) {
    this.to = to;
    this.sendText = sendText;
    this.editText = editText;
    this.intervalMs = intervalMs;
    this.editIntervalMs = editIntervalMs;
    this.setIntervalImpl = setIntervalImpl;
    this.clearIntervalImpl = clearIntervalImpl;
    this.messageKey = null;
    this.completed = [];
    this.active = null;
    this.dots = 1;
    this.timer = null;
    this.editTask = null;
    this.dirty = false;
    this.closed = false;
    this.warned = false;
  }

  async start() {
    if (this.closed || this.messageKey) return;
    if (!this.active) this.active = "Memproses permintaan";
    const message = await this.sendText({ to: this.to, text: this.render() });
    this.messageKey = message?.key ?? null;
    if (!this.messageKey) return;
    this.lastEditAt = Date.now();
    this.timer = this.setIntervalImpl(() => {
      this.dots = this.dots % 3 + 1;
      void this.requestEdit();
    }, this.intervalMs);
    this.timer?.unref?.();
  }

  async update(event) {
    if (!event || this.closed) return;
    if (event.type === "step_start" && typeof event.label === "string") {
      if (!this.messageKey) {
        this.active = event.label;
        this.dots = 1;
        await this.start();
        return;
      }
      if (this.active) this.completed.push({ label: this.active, status: "done" });
      this.active = event.label;
      this.dots = 1;
      await this.requestEdit();
      return;
    }
    if (event.type === "step_end" && this.active) {
      this.completed.push({ label: this.active, status: event.success ? "done" : "failed" });
      this.active = null;
      this.dots = 1;
      await this.requestEdit();
    }
  }

  async complete() {
    if (this.closed) return;
    this.closed = true;
    this.stopTimer();
    if (this.active) this.completed.push({ label: this.active, status: "done" });
    this.active = null;
    if (this.completed.length === 0) this.completed.push({ label: "Memproses permintaan", status: "done" });
    this.completed.push({ label: "Selesai", status: "done" });
    await this.requestEdit();
    await this.editTask;
  }

  async fail() {
    if (this.closed) return;
    this.closed = true;
    this.stopTimer();
    if (this.active) this.completed.push({ label: this.active, status: "failed" });
    this.active = null;
    this.completed.push({ label: "Terjadi kendala", status: "failed" });
    await this.requestEdit();
    await this.editTask;
  }

  render() {
    // Collapse runs of the same label: a 12-round tool loop otherwise prints
    // "Menjalankan langkah agent" a dozen times and pushes real progress off
    // the message. Keep the newest 14 entries so the text stays readable.
    const merged = [];
    for (const { label, status } of this.completed) {
      const previous = merged[merged.length - 1];
      if (previous && previous.label === label && previous.status === status) {
        previous.count = (previous.count ?? 1) + 1;
        continue;
      }
      merged.push({ label, status, count: 1 });
    }
    const lines = merged.slice(-14).map(({ label, status, count }) =>
      `${status === "done" ? ">" : "!"} ${label}${count > 1 ? ` (${count}x)` : ""}`);
    if (this.active) lines.push(`${this.active}${" .".repeat(this.dots)}`);
    if (lines.length === 0) lines.push(`Memproses permintaan${" .".repeat(this.dots)}`);
    return ["Progress", ...lines].join("\n").slice(0, 4000);
  }

  stopTimer() {
    if (this.timer) this.clearIntervalImpl(this.timer);
    this.timer = null;
  }

  requestEdit() {
    if (!this.messageKey || !this.editText) return Promise.resolve();
    this.dirty = true;
    if (this.editTask) return this.editTask;
    this.editTask = (async () => {
      while (this.dirty && this.messageKey) {
        const remaining = this.editIntervalMs - (Date.now() - (this.lastEditAt ?? 0));
        if (remaining > 0) await new Promise((resolve) => setTimeout(resolve, remaining));
        if (!this.dirty || !this.messageKey) break;
        this.dirty = false;
        try {
          await this.editText({ to: this.to, messageKey: this.messageKey, text: this.render() });
          this.lastEditAt = Date.now();
        } catch (error) {
          if (!this.warned) {
            this.warned = true;
            console.warn(`Could not update WhatsApp progress message: ${error.message}`);
          }
          this.messageKey = null;
        }
      }
    })().finally(() => {
      this.editTask = null;
      if (this.dirty) void this.requestEdit();
    });
    return this.editTask;
  }
}