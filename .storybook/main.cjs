const { dirname, join } = require("path");

module.exports = {
    stories: ["../**/src/**/*.stories.tsx"],
    addons: [getAbsolutePath("@storybook/addon-docs")],

    typescript: {
        reactDocgen: false,
    },

    async viteFinal(config) {
        const { mergeConfig } = await import("vite");
        const wyw = await import("@wyw-in-js/vite");
        return mergeConfig(config, {
            plugins: [wyw.default()],
            resolve: {
                alias: [
                    {
                        find: /^@glideapps\/glide-data-grid\/dist\/index\.css$/,
                        replacement: join(__dirname, "source-mode.css"),
                    },
                    {
                        find: /^@glideapps\/glide-data-grid$/,
                        replacement: join(__dirname, "../packages/core/src/index.ts"),
                    },
                ],
            },
        });
    },

    framework: {
        name: getAbsolutePath("@storybook/react-vite"),
        options: {},
    },
};

function getAbsolutePath(value) {
    return dirname(require.resolve(join(value, "package.json")));
}
