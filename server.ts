import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";
import mammoth from "mammoth";
import dotenv from "dotenv";

dotenv.config();

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Increase payload limit to handle larger base64 file uploads
  app.use(express.json({ limit: "50mb" }));
  app.use(express.urlencoded({ limit: "50mb", extended: true }));

  // API Route for Word Doc to JSON conversion
  app.post("/api/parse-docx", async (req, res) => {
    try {
      const { fileData, fileName, extractionType, customPrompt } = req.body;

      if (!fileData) {
        return res.status(400).json({ error: "Ingen fildata skickades." });
      }

      if (!process.env.GEMINI_API_KEY) {
        return res.status(500).json({
          error: "GEMINI_API_KEY är inte konfigurerad på servern. Gå till Settings > Secrets och lägg till din nyckel."
        });
      }

      // Convert Base64 back to Buffer
      const buffer = Buffer.from(fileData, "base64");

      // Extract raw Markdown from docx using mammoth
      let markdownText = "";
      try {
        const result = await (mammoth as any).convertToMarkdown({ buffer });
        markdownText = result.value;
      } catch (err: any) {
        return res.status(400).json({ error: `Kunde inte läsa Word-dokumentet: ${err.message}` });
      }

      if (!markdownText || markdownText.trim().length === 0) {
        return res.status(400).json({ error: "Word-dokumentet verkar vara tomt eller saknar läsbar text." });
      }

      // Initialize Google Gen AI with required telemetry info
      const ai = new GoogleGenAI({
        apiKey: process.env.GEMINI_API_KEY,
        httpOptions: {
          headers: {
            "User-Agent": "aistudio-build",
          }
        }
      });

      // Prepare guidance instructions and schema according to core request
      const systemInstructionPrefix = `Du är en expert på dokumentstruktur och semantisk parsning.
Din uppgift är att analysera hela den medföljande texten från ett Word-dokument (i Markdown-format) och generera en KOMPLETT JSON-fil som representerar ALLT innehåll i dokumentet.

KRAV:
1. Du får absolut INTE sammanfatta eller utesluta sektioner. All text, tabeller och punktlistor från Word-filen ska med i dess helhet.
2. Svara enbart med ett giltigt JSON-objekt enligt schemat. Ingen inledande eller avslutande förklaring.`;

      let systemInstruction = "";
      let responseMimeType = "application/json";
      let responseSchema: any = undefined;

      if (extractionType === "hierarchy") {
        systemInstruction = `${systemInstructionPrefix}
Skapa en hierarkisk trädstruktur av hela dokumentet med alla dess rubriker, paragrafer, punktlistor och tabeller i kronologisk ordning.`;
        responseSchema = {
          type: Type.OBJECT,
          properties: {
            documentTitle: { type: Type.STRING, description: "Dokumentets huvudtitel eller första stora rubrik." },
            sections: {
              type: Type.ARRAY,
              description: "Alla sektioner i dokumentet i den ordning de visas.",
              items: {
                type: Type.OBJECT,
                properties: {
                  heading: { type: Type.STRING, description: "Rubrikens text." },
                  level: { type: Type.INTEGER, description: "Rubriknivå (t.ex. 1 för H1, 2 för H2, 3 för H3)." },
                  blocks: {
                    type: Type.ARRAY,
                    description: "Alla textblock, listor eller tabeller inom denna sektion i läsordning.",
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        type: { type: Type.STRING, description: "Blocktyp: 'paragraph', 'list' eller 'table'." },
                        text: { type: Type.STRING, description: "Textinnehållet. Krävs om blocktyp är 'paragraph'." },
                        listType: { type: Type.STRING, description: "Listtyp: 'ordered' (numrerad) eller 'unordered' (punktlista). Krävs om blocktyp är 'list'." },
                        listItems: {
                          type: Type.ARRAY,
                          items: { type: Type.STRING },
                          description: "Alla punkter i listan. Krävs om blocktyp är 'list'."
                        },
                        tableHeaders: {
                          type: Type.ARRAY,
                          items: { type: Type.STRING },
                          description: "Kolumnnamn för tabellen. Krävs om blocktyp är 'table'."
                        },
                        tableRows: {
                          type: Type.ARRAY,
                          items: {
                            type: Type.ARRAY,
                            items: { type: Type.STRING },
                            description: "En rad i tabellen bestående av cellvärden."
                          },
                          description: "Datamängd för tabellens rader. Krävs om blocktyp är 'table'."
                        }
                      },
                      required: ["type"]
                    }
                  }
                },
                required: ["heading", "level", "blocks"]
              }
            }
          },
          required: ["documentTitle", "sections"]
        };
      } else if (extractionType === "requirements") {
        systemInstruction = `${systemInstructionPrefix}
Identifiera och packa upp alla systemkrav, produktkrav, maskinkrav eller standarder som listas i Word-filen.
Skapa ett unikt ID för varje krav (t.ex. REQ-001, REQ-002 eller använd befintliga ID:n om de finns i texten).
Inkludera detaljerad beskrivning, kategori/sektion, prioritet (Hög/Medium/Låg baserat på 'skall' / 'bör' / 'kan'), verifieringsmetod (Test, Inspektion, Analys, Demonstration) samt status.`;
        responseSchema = {
          type: Type.OBJECT,
          properties: {
            documentTitle: { type: Type.STRING, description: "Dokumentets titel." },
            version: { type: Type.STRING, description: "Eventuellt versionsnummer funnet i texten." },
            requirements: {
              type: Type.ARRAY,
              description: "Alla krav som extraherats ur dokumentet i sin helhet.",
              items: {
                type: Type.OBJECT,
                properties: {
                  id: { type: Type.STRING, description: "Unikt ID för kravet (extraherat eller genererat sekventiellt)." },
                  title: { type: Type.STRING, description: "Kort, tydlig rubrik för kravet." },
                  description: { type: Type.STRING, description: "Hela kravtexten med alla detaljer och villkor." },
                  category: { type: Type.STRING, description: "Vilken sektion, maskindel eller kategori kravet tillhör." },
                  priority: { type: Type.STRING, description: "Prioritetsnivå: Hög, Medium, Låg." },
                  verificationMethod: { type: Type.STRING, description: "Verifiering: Test, Inspektion, Analys, Demonstration eller Ej angivet." },
                  status: { type: Type.STRING, description: "Status, t.ex. Nytt, Godkänt, Utkast." }
                },
                required: ["id", "title", "description", "category", "priority"]
              }
            }
          },
          required: ["documentTitle", "requirements"]
        };
      } else if (extractionType === "parameters") {
        systemInstruction = `${systemInstructionPrefix}
Hitta alla inställningar, kalibreringsparametrar, gränsvärden, börvärden och tekniska parametrar som nämns i dokumentet och sammanställ dem till en komplett lista.
Varje parameter ska ha ett namn, ett nominellt värde, enhet (t.ex. mm, °C, bar, kg), tolerans eller intervall samt sammanhang (var i maskinen/processen parametern används).`;
        responseSchema = {
          type: Type.OBJECT,
          properties: {
            documentTitle: { type: Type.STRING, description: "Titel eller maskinnamn." },
            parameters: {
              type: Type.ARRAY,
              description: "Fullständig lista över alla tekniska parametrar funna i dokumentet.",
              items: {
                type: Type.OBJECT,
                properties: {
                  name: { type: Type.STRING, description: "Parameterns namn." },
                  value: { type: Type.STRING, description: "Det nominella eller inställda värdet." },
                  unit: { type: Type.STRING, description: "Mätenhet (t.ex. 'mm', 'bar', 'ms') eller null om det saknas." },
                  tolerance: { type: Type.STRING, description: "Tillåtna toleranser, max/min-gränser (t.ex. '+/- 10%', 'min 1.2 bar')." },
                  context: { type: Type.STRING, description: "Beskrivning av hur och var denna parameter tillämpas." }
                },
                required: ["name", "value"]
              }
            }
          },
          required: ["documentTitle", "parameters"]
        };
      } else {
        // Custom prompt schema
        systemInstruction = `${systemInstructionPrefix}
Analysera dokumentet fullständigt och strukturera dess JSON-innehåll exakt baserat på följande instruktioner från användaren: "${customPrompt || "Skapa en komplett strukturerad JSON av allt innehåll."}".
Säkerställ att all information från Word-dokumentet är intakt och formaterad i lämpliga fält.`;
      }

      const contents = `Här är hela texten från Word-dokumentet i Markdown-format. Extrahera ALLT innehåll och strukturera det som komplett JSON:\n\n${markdownText}`;

      const config: any = {
        systemInstruction,
        responseMimeType,
      };

      if (responseSchema) {
        config.responseSchema = responseSchema;
      }

      // Generate structured output using gemini-3.5-flash
      const response = await ai.models.generateContent({
        model: "gemini-3.5-flash",
        contents,
        config,
      });

      const responseText = response.text;
      if (!responseText) {
        return res.status(500).json({ error: "Kunde inte generera ett svar från AI-modellen." });
      }

      // Safe JSON Parse
      let parsedJson;
      try {
        parsedJson = JSON.parse(responseText.trim());
      } catch (parseErr) {
        // fallback in case markdown wrapping slipped in
        const cleanedText = responseText.replace(/^```json\s*/i, "").replace(/```\s*$/, "").trim();
        parsedJson = JSON.parse(cleanedText);
      }

      res.json({
        success: true,
        fileName,
        extractionType,
        textLength: markdownText.length,
        rawMarkdown: markdownText,
        result: parsedJson
      });

    } catch (err: any) {
      console.error("Fel vid parsning:", err);
      res.status(500).json({ error: err.message || "Ett internt serverfel inträffade vid behandling av filen." });
    }
  });

  // Health-check API
  app.get("/api/health", (req, res) => {
    res.json({ status: "healthy", time: new Date().toISOString() });
  });

  // Vite middle-ware layout
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`[Server] igång på http://0.0.0.0:${PORT} under NODE_ENV=${process.env.NODE_ENV}`);
  });
}

startServer();
