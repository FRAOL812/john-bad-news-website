import { useState, useEffect, type FormEvent } from "react";

export default function Admin() {
  const [config, setConfig] = useState({
    telebirrNumber: "",
    telebirrAccountName: "",
    paypalDisplayName: "",
    paypalUsername: "",
  });
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState({ type: "", text: "" });

  const webhookUrl = window.__RUNTIME_CONFIG__?.SPREADSHEET_WEBHOOK_URL || "";

  useEffect(() => {
    fetchConfig();
  }, []);

  const fetchConfig = async () => {
    try {
      const response = await fetch(`${webhookUrl}?path=config`);
      const data = await response.json();
      if (data.ok) {
        setConfig({
          telebirrNumber: data.TELEBIRR_NUMBER || "",
          telebirrAccountName: data.TELEBIRR_ACCOUNT_NAME || "",
          paypalDisplayName: data.PAYPAL_DISPLAY_NAME || "",
          paypalUsername: data.PAYPAL_USERNAME || "",
        });
      } else {
        setMessage({ type: "error", text: "Failed to load configuration" });
      }
    } catch (error) {
      setMessage({ type: "error", text: "Network error loading configuration" });
    } finally {
      setLoading(false);
    }
  };

  const handleSubmit = async (e: FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!password) {
      setMessage({ type: "error", text: "Please enter admin password" });
      return;
    }

    setSaving(true);
    setMessage({ type: "", text: "" });

    try {
      const response = await fetch(`${webhookUrl}?path=config`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          password,
          telebirrNumber: config.telebirrNumber,
          telebirrAccountName: config.telebirrAccountName,
          paypalDisplayName: config.paypalDisplayName,
          paypalUsername: config.paypalUsername,
        }),
      });

      const data = await response.json();
      if (data.ok) {
        setMessage({ type: "success", text: "Configuration updated successfully!" });
        setPassword("");
      } else {
        setMessage({ type: "error", text: data.errors?.join(", ") || "Failed to update configuration" });
      }
    } catch (error) {
      setMessage({ type: "error", text: "Network error updating configuration" });
    } finally {
      setSaving(false);
    }
  };

  if (loading) {
    return (
      <div className="page-shell">
        <div className="contact-card">
          <h2>Admin Configuration</h2>
          <p>Loading configuration...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="page-shell">
      <div className="contact-card">
        <h2>Admin Configuration</h2>
        <p>Update payment account details for Telebirr and PayPal</p>

        {message.text && (
          <div className={`notice-box ${message.type === "error" ? "error" : "success"}`}>
            <span>{message.text}</span>
          </div>
        )}

        <form onSubmit={handleSubmit} className="space-y-4">
          <label>
            <span>Admin Password</span>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="Enter admin password"
              required
            />
          </label>

          <div className="section-divider">
            <strong>Telebirr Configuration</strong>
          </div>

          <label>
            <span>Telebirr Phone Number</span>
            <input
              type="text"
              value={config.telebirrNumber}
              onChange={(e) => setConfig({ ...config, telebirrNumber: e.target.value })}
              placeholder="0913885322"
              required
            />
          </label>

          <label>
            <span>Telebirr Account Name</span>
            <input
              type="text"
              value={config.telebirrAccountName}
              onChange={(e) => setConfig({ ...config, telebirrAccountName: e.target.value })}
              placeholder="Fraol Eshetu Hailu"
              required
            />
          </label>

          <div className="section-divider">
            <strong>PayPal Configuration</strong>
          </div>

          <label>
            <span>PayPal Display Name</span>
            <input
              type="text"
              value={config.paypalDisplayName}
              onChange={(e) => setConfig({ ...config, paypalDisplayName: e.target.value })}
              placeholder="Yonatan Woldegiorgis"
              required
            />
          </label>

          <label>
            <span>PayPal Username</span>
            <input
              type="text"
              value={config.paypalUsername}
              onChange={(e) => setConfig({ ...config, paypalUsername: e.target.value })}
              placeholder="@YonatanWoldegiorgis9"
              required
            />
          </label>

          <button type="submit" disabled={saving} className="submit-button">
            {saving ? "Saving..." : "Save Configuration"}
          </button>
        </form>
      </div>
    </div>
  );
}
