const { FusesPlugin } = require('@electron-forge/plugin-fuses');
const { FuseV1Options, FuseVersion } = require('@electron/fuses');

// Distribution is Windows-first. Keep packaged output focused on the
// installer/update path that the app actually uses in production.
const WINDOWS_SETUP_EXE_NAME = 'WorldShotLogSetup.exe';
const WINDOWS_EXECUTABLE_NAME = 'WorldShotLog';
const WINDOWS_ICON_PATH = './img/logo.ico';

// Exclude local caches, VCS metadata, and recovery artifacts from packaged
// builds so release output stays lean and predictable.
const PACKAGER_IGNORE_PATTERNS = [
  /^\/\.git($|\/)/,
  /^\/\.github($|\/)/,
  /^\/\.gitignore$/,
  /^\/assets($|\/)/,
  /^\/docs($|\/)/,
  /^\/index\.html$/,
  /^\/release-notes($|\/)/,
  /^\/scripts($|\/)/,
  /^\/test_pic($|\/)/,
  /^\/tmp($|\/)/,
  /^\/thumbnails($|\/)/,
  /^\/out($|\/)/,
  /^\/make($|\/)/,
  /^\/img\/\.png$/,
  /^\/img\/app_view($|\/)/,
  /^\/img\/booth_logo_icon\.svg$/,
  /^\/img\/editingtools($|\/)/,
  /^\/img\/notuse\.png$/,
  /^\/img\/usecase($|\/)/,
  /^\/img\/vrc($|\/)/,
  /^\/src\/renderer\.js\.bak-corrupt$/,
  /^\/node_modules\/better-sqlite3\/prebuilds\/(?!win32-x64\.node$)/,
];

module.exports = {
  packagerConfig: {
    asar: true,
    prune: true,
    junk: true,
    ignore: PACKAGER_IGNORE_PATTERNS,
    executableName: WINDOWS_EXECUTABLE_NAME,
    icon: WINDOWS_ICON_PATH,
    extraResource: ['./resources/models', './resources/onnxruntime-web'],
    win32metadata: {
      CompanyName: 'nooma',
      FileDescription: 'WorldShot Log',
      InternalName: WINDOWS_EXECUTABLE_NAME,
      OriginalFilename: `${WINDOWS_EXECUTABLE_NAME}.exe`,
      ProductName: 'WorldShot Log',
    },
  },
  // better-sqlite3 13 ships a portable N-API prebuild; rebuilding it for an
  // Electron ABI is unnecessary and would replace the reviewed binary.
  rebuildConfig: { ignoreModules: ['better-sqlite3'] },
  makers: [
    {
      name: '@electron-forge/maker-squirrel',
      platforms: ['win32'],
      config: {
        noMsi: true,
        setupExe: WINDOWS_SETUP_EXE_NAME,
        setupIcon: WINDOWS_ICON_PATH,
        additionalFiles: [
          { src: 'LICENSES.chromium.html', target: 'lib\\net45\\LICENSES.chromium.html' },
          { src: 'version', target: 'lib\\net45\\version' },
        ],
      },
    },
  ],
  plugins: [
    {
      name: '@electron-forge/plugin-auto-unpack-natives',
      config: {},
    },
    // Fuses are used to disable unused Electron behaviors before release.
    new FusesPlugin({
      version: FuseVersion.V1,
      [FuseV1Options.RunAsNode]: false,
      [FuseV1Options.EnableCookieEncryption]: true,
      [FuseV1Options.EnableNodeOptionsEnvironmentVariable]: false,
      [FuseV1Options.EnableNodeCliInspectArguments]: false,
      [FuseV1Options.EnableEmbeddedAsarIntegrityValidation]: true,
      [FuseV1Options.OnlyLoadAppFromAsar]: true,
    }),
  ],
};
