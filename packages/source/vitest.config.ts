import react from "@vitejs/plugin-react";
import { defineConfig, configDefaults } from "vitest/config";
import { fileURLToPath } from "node:url";

export default defineConfig({
    plugins: [react()],
    resolve: {
        alias: {
            "@glideapps/glide-data-grid": fileURLToPath(new URL("../core/src/index.ts", import.meta.url)),
        },
    },
    test: {
        include: ["test/**/*.test.tsx", "test/**/*.test.ts"],
        environment: "jsdom",
        setupFiles: "vitest.setup.ts",
        threads: false,
        singleThread: true,
        watch: false,
        clearMocks: true,
        maxConcurrency: 5,
        fakeTimers: {
            toFake: [
                ...(configDefaults.fakeTimers.toFake ?? []),
                "performance",
                "requestAnimationFrame",
                "cancelAnimationFrame",
            ],
        },
        deps: {
            optimizer: {
                web: {
                    include: ["vitest-canvas-mock"],
                },
            },
        },
        environmentOptions: {
            jsdom: {
                resources: "usable",
            },
        },
    },
});
