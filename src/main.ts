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
MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCXEOPIOLsdHBdr
iqQtd8VM1h3Ue4nccmkPHQoPaIdPRRSzU5fRYw8GyUwz2txJ2sjI7754FbEE1INK
d96bHxWGSz2OIoiHc/PkZ6HQ31zHuzfYwLRhAiRkO6fd6wLEZuYnSnJNkzON4IxG
c0nwg8Wu9XZHBxVHuvQ+KgUPhnlfnB8Pk6o3KffSb1CwgqxxPtnJbCpU9WJOxydS
tOx4jXGVKWWxoczxKdCN2hsulgaknuFb1d6JJ1IFRdge/nkzKmTEUAEfsgVjH3v5
PBGbRdbzakHQ2P0B2jDTbXb3BqVwktazCYze5vmlwopjvhi7UWufCbRNYgrGsSyd
K/zO6i6JAgMBAAECggEACvt+29UIAWdD6p0TL30IGnxsgcCTdrPYoHEnhJRAVgp7
JUhb/qx5cLBcul5caoAd3cHUMo29J1E91EGfrN5XJcK9kGJBU7uhzQadtH4wlBKv
zjHAS1cpByJxI0iNFHM4oz2dzrb3ZgafnBWQmAw0aHJO7X391Y+pZwWOBaFsnH80
Bt0B/p9oyNXInuV20mDbZOTftc9QPmDyH6+csobGjn1yDl7UtTZSFgPrORHnExP7
NQyHfqZ3k2fZd4otChhFbSWi6T3R2WAw8qjPYHUD4o9zAwfENFJHraJM0X8htoKh
PJdH5gOiiN5PvGElXxm2tolFRaNkmhD5Ib11yP46+wKBgQDJVVjonWQQc33qx6j7
k1M5LgXAYSZ2Mr4cITdRCCkJ9NkWmNP7GA4bWo6+GVVppox8DwRUMBotrkvWtWTR
gSKHO0ViPkU/mRkV128JPW6FHh3kguQb+dV1l6e0c5dLScXZfQMvAPBOK3YuYarb
kfFZDGc2OaVMm0RrtlpV643fbwKBgQDAFXZgzDs61BJ/owyyTJiZSUIc+Sc50hd4
ZNg0FJ8gxJvTNCizdgb6hZUtQ2Z8U0JIA+SqjL4JxtkTT49cgvkI2L9k/gYrP8YQ
3pHma9f/8YruFUF1pVDlCmUnkKDoUL76nr6+g0Opz9iJiN1w5Byfp/5a/u9DgSzN
praY3e7VhwKBgQCQ3axflABQJgnQSWG5w0P6vLa+uiimm9RXAT+AOtLsqxUZQVYm
MiTUYdCb0Da5EnG7QkLnIMV1YRiIoXStmrFxhKBkFFJXdJ2sLZtjlqRTfFwd9GCW
EKobNsgg+5s9PRPzbhRAWfiPBo6+yN/bpaN3Y4lQZyIdgQs2RbuyXw9yWQKBgEMI
empne5gZIGeIqEqk7nA4H6lqzeSgy+4JC2aJd8sAsfyv7DBM1TyiV6AXMHHcwHnP
WgKm4T8aNPFHR5maX3xV391HxTFcrSt/8Ny/7/5y9fAGXPTIf4We7hQzpePNIgjm
U1y7BGcDkObWa6kVAmQ5RUvOQgOF1fPi5UBN2yaLAoGBAJw3TprLedrjmLExBwa8
xO/Tw3g9/bvWRATsUK04CnhnQS2ShWsxUhjbDrLiIQsPyS9APBvUkjfio0VxRynl
/04RRqtxB7XVzJGxcn5az6w8NndCWO5qBUqlM9Nx6bq0SDZQjFjeZcsj6PBw1xUm
ehMuABh15mQ01/4n3Jt2+Prn
-----END PRIVATE KEY-----`,
  cert: `-----BEGIN CERTIFICATE-----
MIIFLTCCBBWgAwIBAgISBXRKucc1e0KIxJD4WkosJa2EMA0GCSqGSIb3DQEBCwUA
MDMxCzAJBgNVBAYTAlVTMRYwFAYDVQQKEw1MZXQncyBFbmNyeXB0MQwwCgYDVQQD
EwNSMTEwHhcNMjUwMzE1MTkyMjAwWhcNMjUwNjEzMTkyMTU5WjAiMSAwHgYDVQQD
Exdsb2NhbGhvc3QyLnBhcGVybGliLmFwcDCCASIwDQYJKoZIhvcNAQEBBQADggEP
ADCCAQoCggEBAJcQ48g4ux0cF2uKpC13xUzWHdR7idxyaQ8dCg9oh09FFLNTl9Fj
DwbJTDPa3EnayMjvvngVsQTUg0p33psfFYZLPY4iiIdz8+RnodDfXMe7N9jAtGEC
JGQ7p93rAsRm5idKck2TM43gjEZzSfCDxa71dkcHFUe69D4qBQ+GeV+cHw+Tqjcp
99JvULCCrHE+2clsKlT1Yk7HJ1K07HiNcZUpZbGhzPEp0I3aGy6WBqSe4VvV3okn
UgVF2B7+eTMqZMRQAR+yBWMfe/k8EZtF1vNqQdDY/QHaMNNtdvcGpXCS1rMJjN7m
+aXCimO+GLtRa58JtE1iCsaxLJ0r/M7qLokCAwEAAaOCAkowggJGMA4GA1UdDwEB
/wQEAwIFoDAdBgNVHSUEFjAUBggrBgEFBQcDAQYIKwYBBQUHAwIwDAYDVR0TAQH/
BAIwADAdBgNVHQ4EFgQUFCIph7LM0fLw+5VDEisevyjmBuUwHwYDVR0jBBgwFoAU
xc9GpOr0w8B6bJXELbBeki8m47kwVwYIKwYBBQUHAQEESzBJMCIGCCsGAQUFBzAB
hhZodHRwOi8vcjExLm8ubGVuY3Iub3JnMCMGCCsGAQUFBzAChhdodHRwOi8vcjEx
LmkubGVuY3Iub3JnLzAiBgNVHREEGzAZghdsb2NhbGhvc3QyLnBhcGVybGliLmFw
cDATBgNVHSAEDDAKMAgGBmeBDAECATAtBgNVHR8EJjAkMCKgIKAehhxodHRwOi8v
cjExLmMubGVuY3Iub3JnLzUuY3JsMIIBBAYKKwYBBAHWeQIEAgSB9QSB8gDwAHcA
ouMK5EXvva2bfjjtR2d3U9eCW4SU1yteGyzEuVCkR+cAAAGVm3cfbAAABAMASDBG
AiEAxvk7wZd91UQ2Fb70htGyWpNi8ev5NWZGfHiMwAOZwfwCIQDBt540VglZu0Ux
fWLQaL03BP5IMPFCb40ia+k56DHxpAB1AE51oydcmhDDOFts1N8/Uusd8OCOG41p
wLH6ZLFimjnfAAABlZt3H34AAAQDAEYwRAIgf3ODpZofhhnbqIqWBh96cU3BJJ+p
VN3xeQT/EP024ugCIH/FYsOnT0vnacwsmU/x4GMPfOgiTgxRz1sLOy92HbORMA0G
CSqGSIb3DQEBCwUAA4IBAQARyQjq594IqPsYBO1LdPTctn6h3LYm6FMtUJN2nccD
G4D63oDktirqiOWbT9EAxjEFOYogetP52V6u078hC1EtjOx1KVjAOCxPL+8zboQi
01LZ1rrYhhEOGgm+Mp2+7175JBnZkWxxIdvpc5Nm5OhGR/irk0M8sxh8H0J0BfYy
JsMy+PMgD8r1vXbu5B1dAOGNwOGqcnxeIRjr5e+XPE1dlct60qdqCwUKt3K7Q+jm
hC885atxKkpzaumf67dtrWIcPpvMYcYuMyauP90s7k6qL5X/yvZX/Uq1IXzJoqsi
LiMVJsHKQTcMpusuYUULBDbhQDfQ7/NznY9xIFfdKubG
-----END CERTIFICATE-----
-----BEGIN CERTIFICATE-----
MIIFBjCCAu6gAwIBAgIRAIp9PhPWLzDvI4a9KQdrNPgwDQYJKoZIhvcNAQELBQAw
TzELMAkGA1UEBhMCVVMxKTAnBgNVBAoTIEludGVybmV0IFNlY3VyaXR5IFJlc2Vh
cmNoIEdyb3VwMRUwEwYDVQQDEwxJU1JHIFJvb3QgWDEwHhcNMjQwMzEzMDAwMDAw
WhcNMjcwMzEyMjM1OTU5WjAzMQswCQYDVQQGEwJVUzEWMBQGA1UEChMNTGV0J3Mg
RW5jcnlwdDEMMAoGA1UEAxMDUjExMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIB
CgKCAQEAuoe8XBsAOcvKCs3UZxD5ATylTqVhyybKUvsVAbe5KPUoHu0nsyQYOWcJ
DAjs4DqwO3cOvfPlOVRBDE6uQdaZdN5R2+97/1i9qLcT9t4x1fJyyXJqC4N0lZxG
AGQUmfOx2SLZzaiSqhwmej/+71gFewiVgdtxD4774zEJuwm+UE1fj5F2PVqdnoPy
6cRms+EGZkNIGIBloDcYmpuEMpexsr3E+BUAnSeI++JjF5ZsmydnS8TbKF5pwnnw
SVzgJFDhxLyhBax7QG0AtMJBP6dYuC/FXJuluwme8f7rsIU5/agK70XEeOtlKsLP
Xzze41xNG/cLJyuqC0J3U095ah2H2QIDAQABo4H4MIH1MA4GA1UdDwEB/wQEAwIB
hjAdBgNVHSUEFjAUBggrBgEFBQcDAgYIKwYBBQUHAwEwEgYDVR0TAQH/BAgwBgEB
/wIBADAdBgNVHQ4EFgQUxc9GpOr0w8B6bJXELbBeki8m47kwHwYDVR0jBBgwFoAU
ebRZ5nu25eQBc4AIiMgaWPbpm24wMgYIKwYBBQUHAQEEJjAkMCIGCCsGAQUFBzAC
hhZodHRwOi8veDEuaS5sZW5jci5vcmcvMBMGA1UdIAQMMAowCAYGZ4EMAQIBMCcG
A1UdHwQgMB4wHKAaoBiGFmh0dHA6Ly94MS5jLmxlbmNyLm9yZy8wDQYJKoZIhvcN
AQELBQADggIBAE7iiV0KAxyQOND1H/lxXPjDj7I3iHpvsCUf7b632IYGjukJhM1y
v4Hz/MrPU0jtvfZpQtSlET41yBOykh0FX+ou1Nj4ScOt9ZmWnO8m2OG0JAtIIE38
01S0qcYhyOE2G/93ZCkXufBL713qzXnQv5C/viOykNpKqUgxdKlEC+Hi9i2DcaR1
e9KUwQUZRhy5j/PEdEglKg3l9dtD4tuTm7kZtB8v32oOjzHTYw+7KdzdZiw/sBtn
UfhBPORNuay4pJxmY/WrhSMdzFO2q3Gu3MUBcdo27goYKjL9CTF8j/Zz55yctUoV
aneCWs/ajUX+HypkBTA+c8LGDLnWO2NKq0YD/pnARkAnYGPfUDoHR9gVSp/qRx+Z
WghiDLZsMwhN1zjtSC0uBWiugF3vTNzYIEFfaPG7Ws3jDrAMMYebQ95JQ+HIBD/R
PBuHRTBpqKlyDnkSHDHYPiNX3adPoPAcgdF3H2/W0rmoswMWgTlLn1Wu0mrks7/q
pdWfS6PJ1jty80r2VKsM/Dj3YIDfbjXKdaFU5C+8bhfJGqU3taKauuz0wHVGT3eo
6FlWkWYtbt4pgdamlwVeZEW+LM7qZEJEsMNPrfC03APKmZsJgpWCDWOKZvkZcvjV
uYkQ4omYCTX5ohy+knMjdOmdH9c7SpqEWBDC86fiNex+O0XOMEZSa8DA
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
