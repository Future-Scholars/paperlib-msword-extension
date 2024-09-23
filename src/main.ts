import { PLAPI, PLExtAPI, PLExtension } from "paperlib-api/api";

import fs from "fs";
import { createServer, Server } from "https";
import os from "os";
import path from "path";

import sudo from "sudo-prompt";
import { WebSocket, WebSocketServer } from "ws";
import { PaperFilterOptions } from "./utils/filter-option";

const certs = {
  key: `-----BEGIN PRIVATE KEY-----
MIGHAgEAMBMGByqGSM49AgEGCCqGSM49AwEHBG0wawIBAQQg1M/qC6afapgJXG91
Gx6Z94V2KaM0j8SI3KyzMC7zzBShRANCAATBcejs3Rifm6gxWZru0eSroCkCee0q
5+OxzTjbV4ROPt9M3NIsQB2yYIRBsuhqc0lNc6IXRGIFgWhy9FaonZnS
-----END PRIVATE KEY-----`,
  cert: `-----BEGIN CERTIFICATE-----
MIIDkDCCAxWgAwIBAgISA0nkM1ORzNjhoiXoag0LyBdgMAoGCCqGSM49BAMDMDIx
CzAJBgNVBAYTAlVTMRYwFAYDVQQKEw1MZXQncyBFbmNyeXB0MQswCQYDVQQDEwJF
NTAeFw0yNDA5MjMwOTAxMzZaFw0yNDEyMjIwOTAxMzVaMCIxIDAeBgNVBAMTF2xv
Y2FsaG9zdDIucGFwZXJsaWIuYXBwMFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAE
wXHo7N0Yn5uoMVma7tHkq6ApAnntKufjsc0421eETj7fTNzSLEAdsmCEQbLoanNJ
TXOiF0RiBYFocvRWqJ2Z0qOCAhkwggIVMA4GA1UdDwEB/wQEAwIHgDAdBgNVHSUE
FjAUBggrBgEFBQcDAQYIKwYBBQUHAwIwDAYDVR0TAQH/BAIwADAdBgNVHQ4EFgQU
h7pQYYM0jakKuBHiJjbXuIWha98wHwYDVR0jBBgwFoAUnytfzzwhT50Et+0rLMTG
cIvS1w0wVQYIKwYBBQUHAQEESTBHMCEGCCsGAQUFBzABhhVodHRwOi8vZTUuby5s
ZW5jci5vcmcwIgYIKwYBBQUHMAKGFmh0dHA6Ly9lNS5pLmxlbmNyLm9yZy8wIgYD
VR0RBBswGYIXbG9jYWxob3N0Mi5wYXBlcmxpYi5hcHAwEwYDVR0gBAwwCjAIBgZn
gQwBAgEwggEEBgorBgEEAdZ5AgQCBIH1BIHyAPAAdgA/F0tP1yJHWJQdZRyEvg0S
7ZA3fx+FauvBvyiF7PhkbgAAAZIeUvc+AAAEAwBHMEUCIQCjpNGsL2nL/vqialga
PR4BJUooKMioSoD9c2NzcE05ngIgbwHrKg+hYMKTXUxvC3sbf3DxstVOgssojywq
MMC+YlQAdgBIsONr2qZHNA/lagL6nTDrHFIBy1bdLIHZu7+rOdiEcwAAAZIeUvhs
AAAEAwBHMEUCIEdVx4PajP+Kj6qSM2p/mH7GVpMBW9xX39mQy8c6EmC8AiEAutCE
h/6t3KZRI1StwBLa6oRAc2bdswnqHY6Ow2bcS7AwCgYIKoZIzj0EAwMDaQAwZgIx
ALDogu2mkzqjW/r2DIL3fWHpsyAUkUuSgmYJvK0//VAH414rVNARgb7wVE6nU2om
wwIxAIyPSDcaJzYPY9ieQMc2/g5+Y9EAFp0yS1nOQE84G0KC92vHc6XkzvLzcsbT
lRfSZA==
-----END CERTIFICATE-----
-----BEGIN CERTIFICATE-----
MIIEVzCCAj+gAwIBAgIRAIOPbGPOsTmMYgZigxXJ/d4wDQYJKoZIhvcNAQELBQAw
TzELMAkGA1UEBhMCVVMxKTAnBgNVBAoTIEludGVybmV0IFNlY3VyaXR5IFJlc2Vh
cmNoIEdyb3VwMRUwEwYDVQQDEwxJU1JHIFJvb3QgWDEwHhcNMjQwMzEzMDAwMDAw
WhcNMjcwMzEyMjM1OTU5WjAyMQswCQYDVQQGEwJVUzEWMBQGA1UEChMNTGV0J3Mg
RW5jcnlwdDELMAkGA1UEAxMCRTUwdjAQBgcqhkjOPQIBBgUrgQQAIgNiAAQNCzqK
a2GOtu/cX1jnxkJFVKtj9mZhSAouWXW0gQI3ULc/FnncmOyhKJdyIBwsz9V8UiBO
VHhbhBRrwJCuhezAUUE8Wod/Bk3U/mDR+mwt4X2VEIiiCFQPmRpM5uoKrNijgfgw
gfUwDgYDVR0PAQH/BAQDAgGGMB0GA1UdJQQWMBQGCCsGAQUFBwMCBggrBgEFBQcD
ATASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBSfK1/PPCFPnQS37SssxMZw
i9LXDTAfBgNVHSMEGDAWgBR5tFnme7bl5AFzgAiIyBpY9umbbjAyBggrBgEFBQcB
AQQmMCQwIgYIKwYBBQUHMAKGFmh0dHA6Ly94MS5pLmxlbmNyLm9yZy8wEwYDVR0g
BAwwCjAIBgZngQwBAgEwJwYDVR0fBCAwHjAcoBqgGIYWaHR0cDovL3gxLmMubGVu
Y3Iub3JnLzANBgkqhkiG9w0BAQsFAAOCAgEAH3KdNEVCQdqk0LKyuNImTKdRJY1C
2uw2SJajuhqkyGPY8C+zzsufZ+mgnhnq1A2KVQOSykOEnUbx1cy637rBAihx97r+
bcwbZM6sTDIaEriR/PLk6LKs9Be0uoVxgOKDcpG9svD33J+G9Lcfv1K9luDmSTgG
6XNFIN5vfI5gs/lMPyojEMdIzK9blcl2/1vKxO8WGCcjvsQ1nJ/Pwt8LQZBfOFyV
XP8ubAp/au3dc4EKWG9MO5zcx1qT9+NXRGdVWxGvmBFRAajciMfXME1ZuGmk3/GO
koAM7ZkjZmleyokP1LGzmfJcUd9s7eeu1/9/eg5XlXd/55GtYjAM+C4DG5i7eaNq
cm2F+yxYIPt6cbbtYVNJCGfHWqHEQ4FYStUyFnv8sjyqU8ypgZaNJ9aVcWSICLOI
E1/Qv/7oKsnZCWJ926wU6RqG1OYPGOi1zuABhLw61cuPVDT28nQS/e6z95cJXq0e
K1BcaJ6fJZsmbjRgD5p3mvEf5vdQM7MCEvU0tHbsx2I5mHHJoABHb8KVBgWp/lcX
GWiWaeOyB7RP+OfDtvi2OsapxXiV7vNVs7fMlrRjY1joKaqmmycnBvAq14AEbtyL
sVfOS66B8apkeFX2NY4XPEYV4ZSCe8VHPrdrERk2wILG3T/EGmSIkCYVUMSnjmJd
VQD9F6Na/+zmXCc=
-----END CERTIFICATE-----

`,
};

class PaperlibMSWordExtension extends PLExtension {
  private _httpsServer?: Server;
  private _socketServer?: WebSocketServer;
  private _ws?: WebSocket;

  disposeCallbacks: (() => void)[];

  constructor() {
    super({
      id: "@future-scholars/paperlib-msword-extension",
      defaultPreference: {
        cslDir: {
          type: "pathpicker",
          name: "CSL Directory",
          description: "Import more CSL from this folder.",
          value: "",
          order: 0,
        },
        installed: {
          type: "boolean",
          name: "Installed",
          description:
            "Whether the extension is installed. Toggle to reinstall.",
          value: false,
          order: 1,
        },
      },
    });

    this.disposeCallbacks = [];
  }

  async initialize() {
    try {
      this._httpsServer = createServer(certs);
      this._socketServer = new WebSocketServer({ server: this._httpsServer });
      this._socketServer.on("connection", (ws) => {
        this._ws = ws;
        ws.on("message", this._handler.bind(this));
      });
      this._httpsServer.listen(21993);
    } catch (e) {
      PLAPI.logService.error(
        "Failed to start Paperlib MS Word Extension server.",
        e as Error,
        true,
        "MSWordExt",
      );
    }

    await PLExtAPI.extensionPreferenceService.register(
      this.id,
      this.defaultPreference,
    );

    this.installWordAddin();

    this.disposeCallbacks.push(
      PLExtAPI.extensionPreferenceService.onChanged(
        `${this.id}:installed`,
        (newValue) => {
          if (newValue.value.value === false) {
            PLAPI.logService.info(
              "Reinstalling Word Addin.",
              "",
              false,
              "MSWordExt",
            );
            this.installWordAddin();
          }
        },
      ),
    );

    this.disposeCallbacks.push(() => {
      this._socketServer?.close();
      this._ws?.close();
      this._httpsServer?.close();
    });
  }

  async dispose() {
    for (const disposeCallback of this.disposeCallbacks) {
      disposeCallback();
    }
    PLExtAPI.extensionPreferenceService.unregister(this.id);
  }

  private async _handler(data: string) {
    const message = JSON.parse(data) as {
      type: string;
      params: any;
    };

    switch (message.type) {
      case "search":
        await this._search(message.params);
        break;
      case "csl-names":
        await this._loadCSLNames();
        break;
      case "load-csl":
        await this._loadCSL(message.params);
        break;
    }
  }

  private async _search(params: { query: string }) {
    const result = await PLAPI.paperService.load(
      new PaperFilterOptions({
        search: params.query,
        searchMode: "general",
        limit: 10,
      }).toString(),
      "addTime",
      "desc",
    );

    const responseResult = result.slice(0, 10);
    this._ws?.send(
      JSON.stringify({ type: "search", response: responseResult }),
    );
  }

  private async _loadCSLNames() {
    const cslDir = PLExtAPI.extensionPreferenceService.get(
      this.id,
      "cslDir",
    ) as string;

    if (fs.existsSync(cslDir)) {
      const cslFiles = fs.readdirSync(cslDir);
      const csls = (
        await Promise.all(
          cslFiles.map(async (cslFile) => {
            if (cslFile.endsWith(".csl")) {
              return cslFile.replace(".csl", "");
            } else {
              return "";
            }
          }),
        )
      ).filter((csl) => csl !== "");

      PLAPI.logService.info(
        `Loaded ${csls.length} CSLs.`,
        `${csls}`,
        false,
        "MSWordExt",
      );
      this._ws?.send(JSON.stringify({ type: "csl-names", response: csls }));
    } else {
      this._ws?.send(JSON.stringify({ type: "csl-names", response: [] }));
    }
  }

  private async _loadCSL(name: string) {
    const cslDir = PLExtAPI.extensionPreferenceService.get(
      this.id,
      "cslDir",
    ) as string;
    const cslPath = path.join(cslDir, `${name}.csl`);
    if (fs.existsSync(cslPath)) {
      const csl = fs.readFileSync(cslPath, "utf8");
      this._ws?.send(JSON.stringify({ type: "load-csl", response: csl }));
    } else {
      this._ws?.send(JSON.stringify({ type: "load-csl", response: "" }));
    }
  }

  async installWordAddin() {
    const installed = PLExtAPI.extensionPreferenceService.get(
      this.id,
      "installed",
    ) as boolean;
    if (installed) {
      return;
    }

    PLAPI.logService.info("Installing Word Addin.", "", false, "MSWordExt");

    const manifestUrl =
      "https://paperlib.app/distribution/word_addin/manifest.xml";
    const manifestPath = path.join(os.tmpdir(), "manifest.xml");
    const downloadedPath = await PLExtAPI.networkTool.download(
      manifestUrl,
      manifestPath,
    );

    if (!downloadedPath) {
      PLAPI.logService.error(
        "Failed to download Word Addin manifest.",
        "",
        true,
        "MSWordExt",
      );
      return;
    }

    if (os.platform() === "darwin") {
      fs.mkdirSync(
        path.join(
          os.homedir(),
          "Library/Containers/com.microsoft.Word/Data/Documents/wef/",
        ),
        { recursive: true },
      );
      fs.copyFileSync(
        manifestPath,
        path.join(
          os.homedir(),
          "Library/Containers/com.microsoft.Word/Data/Documents/wef/paperlib.manifest.xml",
        ),
      );
    } else if (os.platform() === "win32") {
      // TODO: Test on win
      const helperUrl =
        "https://distribution.paperlib.app/word_addin/oaloader.exe";
      const helperPath = path.join(os.tmpdir(), "oaloader.exe");
      await PLExtAPI.networkTool.download(helperUrl, helperPath);
      sudo.exec(
        `${helperPath} add ${manifestPath}`,
        { name: "PaperLib" },
        (error, stdout, stderr) => {
          if (error) {
            PLAPI.logService.error(
              "Failed to install Word Addin.",
              error,
              true,
              "MSWordExt",
            );
          }
        },
      );
    } else {
      PLAPI.logService.error(
        "Unsupported platform.",
        os.platform(),
        true,
        "MSWordExt",
      );
    }

    PLExtAPI.extensionPreferenceService.set(this.id, { installed: true });
  }
}

async function initialize() {
  const extension = new PaperlibMSWordExtension();
  await extension.initialize();

  return extension;
}

export { initialize };
