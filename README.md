# Smart Document Engine Templater

This simple tool allows to use Smart Document Engine for filling DOCX template documents using fields recognized from images.

Requires Python 3, python-docx, wxPython, and Smart Document Engine.

Build instructions tested for MacOS, Smart Document Engine v1.11.0, but can be easily adapted for other operating systems supported by Smart Document Engine (Windows, Linux-based).

1. Add your personal signature to the line 95 of `docengine_cli.cpp`, taken from `/doc/README.html` file of Smart Document Engine SDK.

2. Build `docengine_cli.cpp` with a dynamic linking to `libdocengine` (assuming `docengine_cli.cpp` is in `/bin` folder of the SDK:

```
$ clang++ docengine_cli.cpp -O2 -I ../include -L. -l docengine -o docengine_cli -Wl,-rpath,"@executable_path"
```

3. Move `docengine_cli` executable and `libdocengine` binary to the `resources` folder.

4. Move the configuration bundle (from `/data-zip` directory of the SDK) to the `resources` folder.

5. Modify `resources/config.json` if necessary.