# qpdf (static, vendored)

This folder contains a statically built **qpdf 12.4.2** used by
`src/pdf.cpp` to apply and remove standard PDF password protection
(AES-256, revision 6).

| File | What it is |
| --- | --- |
| `include/qpdf/` | qpdf public headers |
| `lib/libqpdf.a` | static qpdf library |
| `lib/libjpeg.a` | static libjpeg (qpdf dependency) |
| `lib/libz.a` | static zlib (qpdf dependency) |
| `LICENSE.txt` | qpdf's Apache License 2.0 |

## How it was built

Built from the official `qpdf-12.4.2.tar.gz` source with MSYS2 UCRT64 g++:

```bat
cmake -S qpdf-12.4.2 -B build -G "MinGW Makefiles" ^
      -DCMAKE_BUILD_TYPE=Release ^
      -DBUILD_SHARED_LIBS=OFF ^
      -DUSE_IMPLICIT_CRYPTO=OFF ^
      -DREQUIRE_CRYPTO_NATIVE=ON ^
      -DSTATIC_JPEG=ON ^
      -DBUILD_DOC=OFF -DINSTALL_MANUAL=OFF -DINSTALL_EXAMPLES=OFF
cmake --build build --parallel 8
```

`REQUIRE_CRYPTO_NATIVE=ON` makes qpdf use its own AES/SHA-2 implementation, so
the tool does not need OpenSSL or GnuTLS at build or run time.

To rebuild the library, repeat the steps above and copy
`build/libqpdf/libqpdf.a` here.
