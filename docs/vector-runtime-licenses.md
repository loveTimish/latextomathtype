# Vector runtime third-party notices

The offline vector sidecar installs production dependencies strictly from `package-lock.json`.
The packaged `node_modules` tree retains each dependency's license file.

| Component | Version | License file | SHA-256 |
|---|---:|---|---|
| MathJax | 3.2.2 | `node_modules/mathjax-full/LICENSE` | `cfc7749b96f63bd31c3c42b5c471bf756814053e847c10f3eb003417bc523d30` |
| Saxon-JS | 2.7.0 | `node_modules/saxon-js/LICENSE.txt` | `73f09f080333cbf539c255dc55ea706a36f31e58ccd2c8f67929a995a768204f` |

Saxon-JS is redistributed unmodified as part of the application. Its license requires the
copyright notice and disclaimer to remain in the distributed documentation/materials; the
original `LICENSE.txt` is therefore retained in the sidecar package.
