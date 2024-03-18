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
MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDc0Htra5fQpPTu
u1x6mIQvmHOLr83qhuHWgtpCcJQD8aKa0hBHcL2tNJezZmu61ivf0UCJWsgiYlLw
ijQH6aJK+p6OlkXGSDq5ZJgg2WlsVv/kF5APW0FSyUXzalhO0t84WbiGYjaF2q05
E4IDHjr3OfHesf0xKsjrJUgfnei8TRy+RM+caU/fZiAVu+W1A2sgGTtiZC8AQOOC
MekQ8x2RvVDAgaSnSYvUF23akEhyUe8WJnTSbBsm2VbnVycMZyz2la/+6lpornN3
g8KHW8YgQXTrIVK3DwiZqogQBq/jod+ugrvJrddzd5poHZ++JTq5JQDetped2YLD
r//EpKxDAgMBAAECggEAHTlEVEupjHYAaoYGb702pVvuUtriDDtssSihPTDMDheR
Nx89A09y8vTmbNpNwKzuopD9kxAeM5rCsk4AI9nyXiz8Bg/yTRMrHnUnQxWzA6Eh
/ax2pumjZBL6PIRjCo+S9lC9gJ+H6sAts8OWrdX25NhY3+m7giHQ9Hn7KSALeLwW
zP/F6S4iiN8qcj2i25mOjYkBNwE3IGyrPQ9kmuC/9Y6bsFe8PNbn8kUckgU2p/lB
rYH16+yhhv9orkfaMX+K5lqx4IsOR/fXumu8AMNWvmWQiCKZtq4csjW665x6alN7
vpzXNycTBnUJBYZ2DSqmrqpZvvIz4vZgQlzjtvHKJQKBgQD/bbZOb43VeiCoMlkM
DMI6soHx7QnFoX0FBU2qcdSimNhLcUUazaww9ALr2RPjA7R1c7UEljqi4JQN7lfo
5a3wx9zBxDy9Esw9ti+99ij/7upDxyse2ymYB1pKBVvj6ZT7twOPlibS8G2o2O1B
T+Wid+GxyLsF1+SXsEBr7GqPZwKBgQDdTvImiROBoGn0V38X5dcyFAaB73ltncGu
HtZpcKqbyLlNLjJLWXSZz67GEF4ybKAcvhC0WEaANdUjBnhBfasZrqr+GfKkG8eB
+Gtd2cfmRM+/tSws8bLH0t2HGLCcSTaMkcjvqnNIDnO6BlPYB4tZlbg8aKQ5YpnQ
Jqkdds/exQKBgC8a1oIEhI2X5inejxlvyOn2PYyWADVYIKwqXDZQo7wQn+LZ0rqs
r1KfzWIdOFOnPUJjwkBETC/5ZpRjHgcvRDKhSQ7a17CupMfEr21C1jDMqJszQbqB
BFyrDnWUI2wWiYkaKSfzstk3yaFXz/k5eMnLfe3BbOwY8mke8eJ1SPmFAoGBAIgT
p89MD+NvqFamii4+lABl0c6JWietjc6rhXkV3sGlPVMYqbItEgYVbki4/cKRii3C
LHFHqinhb+l2a/EQ/WjwPpG5kLmZnyXqgtIVO9X5z6f4FW6ZOy2lGbOc2dNvLQxo
A55iNzpCMKRciadWlDeEWOFEEl56o0sayneke5JlAoGBANP3Eow9VAVVHcLd0Sjp
3lgmQ5W+XgcSQM0eEEKaT1fif7AIWWJfdWFf6VcSll3Wv8OullW4Wg3e9VYCFnFa
2GJVxVGh+2wsTv4kpRbGSSCTextikxCPXZJn8dV4mXwY/3ltfxQwgdon4icDEdSq
A5nL7CSxLkbng2VHyhQnbtwX
-----END PRIVATE KEY-----`,
  cert: `-----BEGIN CERTIFICATE-----
MIIE+zCCA+OgAwIBAgISA4e/d50mVAs6I1LrewPvkBmJMA0GCSqGSIb3DQEBCwUA
MDIxCzAJBgNVBAYTAlVTMRYwFAYDVQQKEw1MZXQncyBFbmNyeXB0MQswCQYDVQQD
EwJSMzAeFw0yNDAzMTgxMDA2MDBaFw0yNDA2MTYxMDA1NTlaMCIxIDAeBgNVBAMT
F2xvY2FsaG9zdDIucGFwZXJsaWIuYXBwMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8A
MIIBCgKCAQEA3NB7a2uX0KT07rtcepiEL5hzi6/N6obh1oLaQnCUA/GimtIQR3C9
rTSXs2ZrutYr39FAiVrIImJS8Io0B+miSvqejpZFxkg6uWSYINlpbFb/5BeQD1tB
UslF82pYTtLfOFm4hmI2hdqtOROCAx469znx3rH9MSrI6yVIH53ovE0cvkTPnGlP
32YgFbvltQNrIBk7YmQvAEDjgjHpEPMdkb1QwIGkp0mL1Bdt2pBIclHvFiZ00mwb
JtlW51cnDGcs9pWv/upaaK5zd4PCh1vGIEF06yFStw8ImaqIEAav46HfroK7ya3X
c3eaaB2fviU6uSUA3raXndmCw6//xKSsQwIDAQABo4ICGTCCAhUwDgYDVR0PAQH/
BAQDAgWgMB0GA1UdJQQWMBQGCCsGAQUFBwMBBggrBgEFBQcDAjAMBgNVHRMBAf8E
AjAAMB0GA1UdDgQWBBQSWvWsAlc4ix+iLfuMZr0tDglpITAfBgNVHSMEGDAWgBQU
LrMXt1hWy65QCUDmH6+dixTCxjBVBggrBgEFBQcBAQRJMEcwIQYIKwYBBQUHMAGG
FWh0dHA6Ly9yMy5vLmxlbmNyLm9yZzAiBggrBgEFBQcwAoYWaHR0cDovL3IzLmku
bGVuY3Iub3JnLzAiBgNVHREEGzAZghdsb2NhbGhvc3QyLnBhcGVybGliLmFwcDAT
BgNVHSAEDDAKMAgGBmeBDAECATCCAQQGCisGAQQB1nkCBAIEgfUEgfIA8AB2ADtT
d3U+LbmAToswWwb+QDtn2E/D9Me9AA0tcm/h+tQXAAABjlE9YHAAAAQDAEcwRQIg
HnG/U5c9GAXcjkiGeCYNfiKG/0clBFjIoNzyvx+vFSwCIQCHSG9A8J4Lc5buexQ3
ibfsuQVea9JNBAjP5jZrLYXEjwB2AEiw42vapkc0D+VqAvqdMOscUgHLVt0sgdm7
v6s52IRzAAABjlE9YmQAAAQDAEcwRQIgGDJXjBCjVgVcjFmjnfu73vs400TbIYWV
6s4rqxUMKIwCIQD39PcHsOEqnfLxlTMP5xUhL1Quu1Tn+w2smixPkVyNCDANBgkq
hkiG9w0BAQsFAAOCAQEAHEHRzszVXdboupypOtma4VqJ/Ahd3xkEz4suqGlaAAWi
+47U+uKgNR6zuLgdB52BS3clyCIWksQp3ZRbJSNlv2bNfmZawuzF8jQRmKHeh/Qa
p078cRutNMKqRs9dpkuXnvpLbYasgmpLXtJV9Fy24iFFponyVGgURPMFlC/eUgiz
EPJoTeqANXcj377fuQPrgkWvsa1S13Qz9tuZ19ISneghYCTUXZskIazhCL/+2ZP2
H1LK9XwJxMuYZVZ/q3rKI6CfuwWIWtiZEPS4EeV+d44SXQyQS2jqUKycyPPxbn2S
CEP4zW4e18ttFOHDDK4YOFZo/cxofZPnypqQq7nb0w==
-----END CERTIFICATE-----
-----BEGIN CERTIFICATE-----
MIIFFjCCAv6gAwIBAgIRAJErCErPDBinU/bWLiWnX1owDQYJKoZIhvcNAQELBQAw
TzELMAkGA1UEBhMCVVMxKTAnBgNVBAoTIEludGVybmV0IFNlY3VyaXR5IFJlc2Vh
cmNoIEdyb3VwMRUwEwYDVQQDEwxJU1JHIFJvb3QgWDEwHhcNMjAwOTA0MDAwMDAw
WhcNMjUwOTE1MTYwMDAwWjAyMQswCQYDVQQGEwJVUzEWMBQGA1UEChMNTGV0J3Mg
RW5jcnlwdDELMAkGA1UEAxMCUjMwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEK
AoIBAQC7AhUozPaglNMPEuyNVZLD+ILxmaZ6QoinXSaqtSu5xUyxr45r+XXIo9cP
R5QUVTVXjJ6oojkZ9YI8QqlObvU7wy7bjcCwXPNZOOftz2nwWgsbvsCUJCWH+jdx
sxPnHKzhm+/b5DtFUkWWqcFTzjTIUu61ru2P3mBw4qVUq7ZtDpelQDRrK9O8Zutm
NHz6a4uPVymZ+DAXXbpyb/uBxa3Shlg9F8fnCbvxK/eG3MHacV3URuPMrSXBiLxg
Z3Vms/EY96Jc5lP/Ooi2R6X/ExjqmAl3P51T+c8B5fWmcBcUr2Ok/5mzk53cU6cG
/kiFHaFpriV1uxPMUgP17VGhi9sVAgMBAAGjggEIMIIBBDAOBgNVHQ8BAf8EBAMC
AYYwHQYDVR0lBBYwFAYIKwYBBQUHAwIGCCsGAQUFBwMBMBIGA1UdEwEB/wQIMAYB
Af8CAQAwHQYDVR0OBBYEFBQusxe3WFbLrlAJQOYfr52LFMLGMB8GA1UdIwQYMBaA
FHm0WeZ7tuXkAXOACIjIGlj26ZtuMDIGCCsGAQUFBwEBBCYwJDAiBggrBgEFBQcw
AoYWaHR0cDovL3gxLmkubGVuY3Iub3JnLzAnBgNVHR8EIDAeMBygGqAYhhZodHRw
Oi8veDEuYy5sZW5jci5vcmcvMCIGA1UdIAQbMBkwCAYGZ4EMAQIBMA0GCysGAQQB
gt8TAQEBMA0GCSqGSIb3DQEBCwUAA4ICAQCFyk5HPqP3hUSFvNVneLKYY611TR6W
PTNlclQtgaDqw+34IL9fzLdwALduO/ZelN7kIJ+m74uyA+eitRY8kc607TkC53wl
ikfmZW4/RvTZ8M6UK+5UzhK8jCdLuMGYL6KvzXGRSgi3yLgjewQtCPkIVz6D2QQz
CkcheAmCJ8MqyJu5zlzyZMjAvnnAT45tRAxekrsu94sQ4egdRCnbWSDtY7kh+BIm
lJNXoB1lBMEKIq4QDUOXoRgffuDghje1WrG9ML+Hbisq/yFOGwXD9RiX8F6sw6W4
avAuvDszue5L3sz85K+EC4Y/wFVDNvZo4TYXao6Z0f+lQKc0t8DQYzk1OXVu8rp2
yJMC6alLbBfODALZvYH7n7do1AZls4I9d1P4jnkDrQoxB3UqQ9hVl3LEKQ73xF1O
yK5GhDDX8oVfGKF5u+decIsH4YaTw7mP3GFxJSqv3+0lUFJoi5Lc5da149p90Ids
hCExroL1+7mryIkXPeFM5TgO9r0rvZaBFOvV2z0gp35Z0+L4WPlbuEjN/lxPFin+
HlUjr8gRsI3qfJOQFy/9rKIJR0Y/8Omwt/8oTWgy1mdeHmmjk7j1nYsvC9JSQ6Zv
MldlTTKB3zhThV1+XWYp6rjd5JW1zbVWEkLNxE7GJThEUG3szgBVGP7pSWTUTsqX
nLRbwHOoq7hHwg==
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
