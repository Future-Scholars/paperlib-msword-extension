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
MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCzjCeNbE1zHT51
Wig5uD3Mzbc6pwUIrlus6bi/O0pifSpFFTZU8TxJCeOIOqXlCyuTXJZtJ4TXeLOb
y7bnegRtkYuIzK3PKL3ciYfWkqmDqOmAbCW0Hy40KTYqoR7IEtEFeR3jMZxfhZAA
VNoqhgeCNxydSP6Jui7nUa8FSgCxLXawc8KLEKnr6VPjNNg/b2cE6IgT94X+czp4
s8bH5p0KBHRRM9uA1VEicAzBf9fsZorU4uV5kIOCVvgAO19Cpl2gZtlNv9tBf+Kr
0vS4vq7wmmuRgdIhF9pQ+sz6iYJhA0zFlwJ2/m8vJ+ilUy64X31qk3hPLHUqE6kd
sLmiwt4rAgMBAAECggEADh9NOFLdpkF7zJMybGFqsWJrS3iG06LGbZiJaLeeWpZp
urrrMrP+zLSznUV8dii/KeOzphMTrQPYOYrecHqPQQCce9gaW7ViqadC2sSSaY6u
cksPWomRT5aZchgjA0dBNZYyaZ2MJSwTCLoKQIn/3f6KIgnAK5xMKcAQOiRF+Ny7
j9cSuQ8ucupka8Mu/eQhRzEsYXT8dp1+TEOqiMP4OBTrGblOKOIVMGMAr07VZSgq
0Jh1UwzmFiI6Yy/7My5IVxhEcVStlVHvsM122P4QhGBSY5De1mc4g9FASX7KTpzo
3C7kt75vX466f2WNhXTgMx5m0XYGJWQ9jccNtxu34QKBgQDGVz6/qWcddPA2GQLG
6it22VSGUSvPAwjqYhH4EckYnJ7D6L8R08q3SYHqFMIITrTUHJG738hZaoLkLbWv
qSFcq+yAcHzVCIY66S7A4udL+XOdXh5WEXWfJYD1M2yGL0mPciExbBz4RKnGUHoj
K8VcBLL/kRFTHT7tDKv8tgfsMwKBgQDnvknNQmO/aeFKng4siNGi+AamJu1FS3QU
ry1RvT3eQRThFcciSMVUBL4q+5bwpuS/j88H7yFw1zEooVF2JuyfloGZ175H5XJs
hqbDN8o9jsJZmVAYohhm66lQo5SzibTrLgHR559q5eYay+BUZ1DOV74q/3WtX1S3
rOiYOd/OKQKBgAkw79J36jReP+dx30QSg/Mc/SLATjRoopgh9U02ncgLMfxII9qS
ovk9aczMK3WxGAYgUMyRATrLicdDKwE56DbgLLSDAfXpUDcYqTb9DNTjeW0YeHVq
l7XJSiGSwXuyY0lHc6xTo0AKBogPIKnSlHHAMf9P3KqqV0kq5ilu0g0rAoGAaxUv
rwwNYWaAduU/8W4rSF3JXL9CBjIOanxjuZBzZR63kiZpBLuRivhCE0R8A6lqq+W8
qZLi5exZx8d7B9iGoFuAeWEKiNhKHkG+DxjZd8Zeod5I94j3M5+TdjKQRMHN+pog
tyiLLm8a+6jXeMjguugqdF3kt38Ee3cHZ0fe1bkCgYEAmMtiET5rHHTrVe2HiHBC
8nFzNSfXhFU2Mqhk2cBHAxbGV9tP3leGjAKLk17I/IUyVCA8x/UbIMJp4M/esPt3
Zc8PKf/tePCSUVD3lhQGcV/YWrvWhS8pOSDB3EAoeh8n16zYF0/8jh9XUQPEFCbn
WgdMF0+ce+mxlhNG+jfDzSI=
-----END PRIVATE KEY-----`,
  cert: `-----BEGIN CERTIFICATE-----
MIIE+jCCA+KgAwIBAgISA3Dihkjqd9RhL76rUN4N3rNvMA0GCSqGSIb3DQEBCwUA
MDIxCzAJBgNVBAYTAlVTMRYwFAYDVQQKEw1MZXQncyBFbmNyeXB0MQswCQYDVQQD
EwJSMzAeFw0yMzEyMjExMTA0MjRaFw0yNDAzMjAxMTA0MjNaMCIxIDAeBgNVBAMT
F2xvY2FsaG9zdDIucGFwZXJsaWIuYXBwMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8A
MIIBCgKCAQEAs4wnjWxNcx0+dVooObg9zM23OqcFCK5brOm4vztKYn0qRRU2VPE8
SQnjiDql5Qsrk1yWbSeE13izm8u253oEbZGLiMytzyi93ImH1pKpg6jpgGwltB8u
NCk2KqEeyBLRBXkd4zGcX4WQAFTaKoYHgjccnUj+ibou51GvBUoAsS12sHPCixCp
6+lT4zTYP29nBOiIE/eF/nM6eLPGx+adCgR0UTPbgNVRInAMwX/X7GaK1OLleZCD
glb4ADtfQqZdoGbZTb/bQX/iq9L0uL6u8JprkYHSIRfaUPrM+omCYQNMxZcCdv5v
LyfopVMuuF99apN4Tyx1KhOpHbC5osLeKwIDAQABo4ICGDCCAhQwDgYDVR0PAQH/
BAQDAgWgMB0GA1UdJQQWMBQGCCsGAQUFBwMBBggrBgEFBQcDAjAMBgNVHRMBAf8E
AjAAMB0GA1UdDgQWBBSo24l9WHAvbJagbUrKkx3MFIFn9TAfBgNVHSMEGDAWgBQU
LrMXt1hWy65QCUDmH6+dixTCxjBVBggrBgEFBQcBAQRJMEcwIQYIKwYBBQUHMAGG
FWh0dHA6Ly9yMy5vLmxlbmNyLm9yZzAiBggrBgEFBQcwAoYWaHR0cDovL3IzLmku
bGVuY3Iub3JnLzAiBgNVHREEGzAZghdsb2NhbGhvc3QyLnBhcGVybGliLmFwcDAT
BgNVHSAEDDAKMAgGBmeBDAECATCCAQMGCisGAQQB1nkCBAIEgfQEgfEA7wB1AEiw
42vapkc0D+VqAvqdMOscUgHLVt0sgdm7v6s52IRzAAABjIxDNicAAAQDAEYwRAIg
P8VkjwmCPRogpP+Z+SQa3l0NXyxSMwjW4qruJJ5o1bsCIHxXSo2UrarPkBFy3i+Y
JYODZsMiqa/TpAqEiJsVhDCCAHYAdv+IPwq2+5VRwmHM9Ye6NLSkzbsp3GhCCp/m
Z0xaOnQAAAGMjEM2gAAABAMARzBFAiEAncEsbL0HRDvzcPsColH1orGdUKkpCR3w
lciAsvaThaICICf3khEkNpdmOwsxisOMkeAuCO7wxOVHVGCPKlsqd1ucMA0GCSqG
SIb3DQEBCwUAA4IBAQCHBFPlg4iW6A9cj1Aoq6lJ6ZOenX/d1RAl14Lbt4Nlm280
l57/cJmNstHgH3TY4PzlnCvHkDSAWCLH4To9yKFWC/sEPvsbO1H1L/8NCw76aPZK
7uBMajeSUGL03Z1AVbQixKm3HIVZWM06CZD1uceBDoguuVX3AaygW0tUqCGr2H7d
MMXSKzprak7gpNywOG0nC+sgcyV4W0it9HBYuNXe1/6WrY1Ed47/Hmt1xXnnbuTm
Wp5mMrKf4SP1k+/pXow1mD37efiMtdVGYoyLnf8plbIhhIHr1gGv02tLJsaE1w3k
i3q5qOt4R6L9/WmG09FhGahwdjY71dCodm+XWM8W
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
-----BEGIN CERTIFICATE-----
MIIFYDCCBEigAwIBAgIQQAF3ITfU6UK47naqPGQKtzANBgkqhkiG9w0BAQsFADA/
MSQwIgYDVQQKExtEaWdpdGFsIFNpZ25hdHVyZSBUcnVzdCBDby4xFzAVBgNVBAMT
DkRTVCBSb290IENBIFgzMB4XDTIxMDEyMDE5MTQwM1oXDTI0MDkzMDE4MTQwM1ow
TzELMAkGA1UEBhMCVVMxKTAnBgNVBAoTIEludGVybmV0IFNlY3VyaXR5IFJlc2Vh
cmNoIEdyb3VwMRUwEwYDVQQDEwxJU1JHIFJvb3QgWDEwggIiMA0GCSqGSIb3DQEB
AQUAA4ICDwAwggIKAoICAQCt6CRz9BQ385ueK1coHIe+3LffOJCMbjzmV6B493XC
ov71am72AE8o295ohmxEk7axY/0UEmu/H9LqMZshftEzPLpI9d1537O4/xLxIZpL
wYqGcWlKZmZsj348cL+tKSIG8+TA5oCu4kuPt5l+lAOf00eXfJlII1PoOK5PCm+D
LtFJV4yAdLbaL9A4jXsDcCEbdfIwPPqPrt3aY6vrFk/CjhFLfs8L6P+1dy70sntK
4EwSJQxwjQMpoOFTJOwT2e4ZvxCzSow/iaNhUd6shweU9GNx7C7ib1uYgeGJXDR5
bHbvO5BieebbpJovJsXQEOEO3tkQjhb7t/eo98flAgeYjzYIlefiN5YNNnWe+w5y
sR2bvAP5SQXYgd0FtCrWQemsAXaVCg/Y39W9Eh81LygXbNKYwagJZHduRze6zqxZ
Xmidf3LWicUGQSk+WT7dJvUkyRGnWqNMQB9GoZm1pzpRboY7nn1ypxIFeFntPlF4
FQsDj43QLwWyPntKHEtzBRL8xurgUBN8Q5N0s8p0544fAQjQMNRbcTa0B7rBMDBc
SLeCO5imfWCKoqMpgsy6vYMEG6KDA0Gh1gXxG8K28Kh8hjtGqEgqiNx2mna/H2ql
PRmP6zjzZN7IKw0KKP/32+IVQtQi0Cdd4Xn+GOdwiK1O5tmLOsbdJ1Fu/7xk9TND
TwIDAQABo4IBRjCCAUIwDwYDVR0TAQH/BAUwAwEB/zAOBgNVHQ8BAf8EBAMCAQYw
SwYIKwYBBQUHAQEEPzA9MDsGCCsGAQUFBzAChi9odHRwOi8vYXBwcy5pZGVudHJ1
c3QuY29tL3Jvb3RzL2RzdHJvb3RjYXgzLnA3YzAfBgNVHSMEGDAWgBTEp7Gkeyxx
+tvhS5B1/8QVYIWJEDBUBgNVHSAETTBLMAgGBmeBDAECATA/BgsrBgEEAYLfEwEB
ATAwMC4GCCsGAQUFBwIBFiJodHRwOi8vY3BzLnJvb3QteDEubGV0c2VuY3J5cHQu
b3JnMDwGA1UdHwQ1MDMwMaAvoC2GK2h0dHA6Ly9jcmwuaWRlbnRydXN0LmNvbS9E
U1RST09UQ0FYM0NSTC5jcmwwHQYDVR0OBBYEFHm0WeZ7tuXkAXOACIjIGlj26Ztu
MA0GCSqGSIb3DQEBCwUAA4IBAQAKcwBslm7/DlLQrt2M51oGrS+o44+/yQoDFVDC
5WxCu2+b9LRPwkSICHXM6webFGJueN7sJ7o5XPWioW5WlHAQU7G75K/QosMrAdSW
9MUgNTP52GE24HGNtLi1qoJFlcDyqSMo59ahy2cI2qBDLKobkx/J3vWraV0T9VuG
WCLKTVXkcGdtwlfFRjlBz4pYg1htmf5X6DYO8A4jqv2Il9DjXA6USbW1FzXSLr9O
he8Y4IWS6wY7bCkjCWDcRQJMEhg76fsO3txE+FiYruq9RUWhiF1myv4Q6W+CyBFC
Dfvp7OOGAN6dEOM4+qR9sdjoSYKEBpsr6GtPAQw4dy753ec5
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
    const cslDir = PLExtAPI.extensionPreferenceService.get(this.id, "cslDir")
      .value as string;

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
    const cslDir = PLExtAPI.extensionPreferenceService.get(this.id, "cslDir")
      .value as string;
    const cslPath = path.join(cslDir, `${name}.csl`);
    if (fs.existsSync(cslPath)) {
      const csl = fs.readFileSync(cslPath, "utf8");
      this._ws?.send(JSON.stringify({ type: "load-csl", response: csl }));
    } else {
      this._ws?.send(JSON.stringify({ type: "load-csl", response: "" }));
    }
  }

  async installWordAddin() {
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
  }
}

async function initialize() {
  const extension = new PaperlibMSWordExtension();
  await extension.initialize();

  return extension;
}

export { initialize };
