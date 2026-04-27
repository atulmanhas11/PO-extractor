import { useState } from "react";

export default function POExtractor() {
  const [input, setInput] = useState("");
  const [output, setOutput] = useState("");

  const handleExtract = async () => {
    const res = await fetch("/api/extract", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: "gpt-4.1",
        input: `Extract structured PO data from this: ${input}`
      })
    });

    const data = await res.json();
    setOutput(JSON.stringify(data, null, 2));
  };

  return (
    <div>
      <textarea
        rows={10}
        style={{ width: "100%" }}
        placeholder="Paste PO text here..."
        value={input}
        onChange={(e) => setInput(e.target.value)}
      />
      <br /><br />
      <button onClick={handleExtract}>Extract</button>

      <pre style={{ marginTop: 20, background: "#f5f5f5", padding: 10 }}>
        {output}
      </pre>
    </div>
  );
}