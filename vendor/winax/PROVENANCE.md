# Vendored winax

Source: https://www.npmjs.com/package/winax/v/3.6.9

Repository: https://github.com/durs/node-activex ; upstream package git head `8cdae1053db066b21240ddbc88a81b72ce199257`.

Original tarball integrity: `sha512-R+6yTIk8pnIf50P4z8unG3yfRP0DCYzq7v8J04tool1OR4UkoXKEGfAj+cq6IPPhnVQHz7+m7isg+yU38T727w==`. Integrity was verified before copying only native runtime source, binding.gyp, index/activex and MIT LICENSE. No examples, Office macro fixture, WScript CLI, tests, binaries or upstream workflows are included.

Local change: `src/utils.h` uses a SFINAE-safe optional `HolderV2()` detector and explicit dependent type qualifiers, retaining `This()` on Node22 and `HolderV2()` on Node24. The upstream3.6.9 alias accessed a missing member before substitution and failed to compile on Node22. Line endings and trailing whitespace in copied source/JS/gyp files were normalized; the semantic change is confined to utils.h. The local package metadata only identifies the vendored CommonJS module.

The root install script builds this exact source with pinned node-gyp13.0.2 on Windows. Real native COM and installed-package CI must pass on both supported Node versions. Remove the local patch only after an upstream version passes those gates.
