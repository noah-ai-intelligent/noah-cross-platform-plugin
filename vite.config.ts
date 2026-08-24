import { defineConfig, ViteDevServer, UserConfig } from "vite";
import react from "@vitejs/plugin-react";
import { resolve } from "path";
import { viteSingleFile } from "vite-plugin-singlefile";

// Office Add-ins require the dev server (and every hosted page) to be served
// over HTTPS, even locally — `office-addin-dev-certs` installs a trusted
// localhost cert once (`npx office-addin-dev-certs install`), which Vite
// picks up here. Three entry points: the taskpane (the main UI), the login
// dialog (opened via Office.context.ui.displayDialogAsync for the OAuth
// flow), and the ribbon-command function file.
export default defineConfig(async (): Promise<UserConfig> => {
  let httpsOptions;
  if (process.env.NODE_ENV !== 'production') {
    const devCerts = await import("office-addin-dev-certs");
    httpsOptions = await devCerts.getHttpsServerOptions();
  }
  
  const isGoogle = process.env.GOOGLE === "true";

  return {
    plugins: [
      react(),
      isGoogle && viteSingleFile(),
      {
        name: "rewrite-middleware",
        configureServer(server: ViteDevServer) {
          server.middlewares.use((req, _res, next) => {
            if (!req.url) return next();
            
            const [path, query] = req.url.split('?');
            const queryString = query ? `?${query}` : '';
            
            if (path === "/taskpane" || path === "/taskpane/") {
              req.url = `/src/taskpane/index.html${queryString}`;
            } else if (path === "/login" || path === "/login/") {
              req.url = `/src/login/login.html${queryString}`;
            } else if (path === "/commands" || path === "/commands/") {
              req.url = `/src/commands/commands.html${queryString}`;
            }
            next();
          });
        },
      },
    ],
    server: {
      port: 5175,
      https: httpsOptions,
    },
    build: { target: "es2015",
      rollupOptions: {
        input: isGoogle
          ? {
              taskpane: resolve(__dirname, "src/taskpane/index.html"),
            }
          : {
              taskpane: resolve(__dirname, "src/taskpane/index.html"),
              login: resolve(__dirname, "src/login/login.html"),
              commands: resolve(__dirname, "src/commands/commands.html"),
            },
      },
    },
  };
});
