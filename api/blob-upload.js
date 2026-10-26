import { handleUpload } from "@vercel/blob/client";

const allowedContentTypes = [
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
  "application/vnd.ms-excel.sheet.macroEnabled.12",
];
const maxFileSize = 20 * 1024 * 1024;

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    res.status(405).json({ error: "Method not allowed." });
    return;
  }

  let body;
  try {
    if (Buffer.isBuffer(req.body)) {
      body = JSON.parse(req.body.toString("utf8"));
    } else if (typeof req.body === "string") {
      body = JSON.parse(req.body);
    } else if (req.body && typeof req.body === "object") {
      body = req.body;
    } else {
      const rawBody = [];
      for await (const chunk of req) rawBody.push(chunk);
      body = JSON.parse(Buffer.concat(rawBody).toString("utf8"));
    }
  } catch {
    res.status(400).json({ error: "Invalid upload request." });
    return;
  }

  if (!process.env.BLOB_READ_WRITE_TOKEN) {
    res.status(503).json({ error: "Connect a Vercel Blob store before uploading workbooks." });
    return;
  }

  try {
    const response = await handleUpload({
      body,
      request: req,
      onBeforeGenerateToken: async (pathname) => {
        if (!pathname.startsWith("uploads/")) {
          throw new Error("Workbooks must be uploaded under the uploads path.");
        }
        const extension = pathname.slice(pathname.lastIndexOf(".")).toLowerCase();
        if (![".xlsx", ".xls", ".xlsm"].includes(extension)) {
          throw new Error("Only Excel workbooks can be uploaded.");
        }

        return {
          allowedContentTypes,
          maximumSizeInBytes: maxFileSize,
          addRandomSuffix: true,
        };
      },
    });

    res.status(200).json(response);
  } catch (error) {
    res.status(400).json({ error: error.message || "Could not authorize workbook upload." });
  }
}
