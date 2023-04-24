
let icon = './img/vega_logo.icns';
if (process.platform !== "darwin") {
  icon = './img/vega_logo.ico';
}

module.exports = {
  packagerConfig: {
    icon: icon, // no file extension required
  },
  rebuildConfig: {},
  makers: [
    {
      name: '@electron-forge/maker-squirrel',
      config: {},
    },
    {
      name: '@electron-forge/maker-zip',
      platforms: ['darwin'],
    },
    {
      name: '@electron-forge/maker-deb',
      config: {},
    },
    {
      name: '@electron-forge/maker-rpm',
      config: {},
    },
  ],
};
