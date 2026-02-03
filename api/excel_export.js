import fetch from "node-fetch";
import * as XLSX from "xlsx";

export default async function handler(req, res) {
  // -----------------------------
  // CORS
  // -----------------------------
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader(
    "Access-Control-Allow-Headers",
    "Content-Type, Authorization"
  );
  res.setHeader("Access-Control-Allow-Credentials", "true");

  // ⚠️ FONTOS: OPTIONS preflight válasz
  if (req.method === "OPTIONS") {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
    res.setHeader(
      "Access-Control-Allow-Headers",
      "Content-Type, Authorization"
    );
    res.setHeader("Access-Control-Allow-Credentials", "true");

    return res.status(204).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  const {
    context,
    statusColumnId,
    allowedStatus,
    targetStatus,
    columnIds,
  } = req.body || {};

  const boardId = context?.boardId;

  const token = process.env.MONDAY_API_KEY;

  if (
    !token ||
    !boardId ||
    !statusColumnId ||
    !allowedStatus ||
    !targetStatus ||
    !Array.isArray(columnIds)
  ) {
    return res.status(400).json({ error: "Hiányzó vagy hibás paraméterek" });
  }

  try {
    const itemsQuery = `
      query {
        boards(ids: ${boardId}) {
          items {
            id
            name
            column_values {
              id
              text
            }
          }
        }
      }
    `;

    const itemsResponse = await fetch("https://api.monday.com/v2", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: token,
      },
      body: JSON.stringify({ query: itemsQuery }),
    });

    const itemsJson = await itemsResponse.json();
    const items = itemsJson?.data?.boards?.[0]?.items || [];

    const filteredItems = items.filter((item) => {
      const statusCol = item.column_values.find(
        (cv) => cv.id === statusColumnId
      );
      return statusCol?.text?.trim() === allowedStatus;
    });

    if (filteredItems.length === 0) {
      return res.status(200).json({
        message: "Nincs exportálható rekord.",
      });
    }

    const rows = filteredItems.map((item) => {
      const row = { Név: item.name };
      for (const col of columnIds) {
        const colValue = item.column_values.find((cv) => cv.id === col.id);
        row[col.label] = colValue?.text ?? "";
      }
      return row;
    });

    const headers = ["Név", ...columnIds.map((c) => c.label)];
    const worksheet = XLSX.utils.json_to_sheet(rows, { header: headers });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Új belépők");

    const buffer = XLSX.write(workbook, {
      type: "buffer",
      bookType: "xlsx",
    });

    for (const item of filteredItems) {
      const mutation = `
        mutation {
          change_column_value(
            board_id: ${boardId},
            item_id: ${item.id},
            column_id: "${statusColumnId}",
            value: "{\\"label\\": \\"${targetStatus}\\"}"
          ) {
            id
          }
        }
      `;

      await fetch("https://api.monday.com/v2", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: token,
        },
        body: JSON.stringify({ query: mutation }),
      });
    }

    const now = new Date()
      .toISOString()
      .replace("T", "_")
      .replace(/:/g, "-")
      .replace(/\..+/, "");

    const filename = `EFO_uj_belepok_${now}.xlsx`;

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${filename}"`
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    return res.status(200).send(buffer);
  } catch (err) {
    console.error("Excel export hiba:", err);
    return res.status(500).json({ error: "Excel export hiba" });
  }
}
