# Docx -> Rerun plugin
This is an example data-loader plugin that lets you view docx files in the [Rerun](https://github.com/rerun-io/rerun/) Viewer.
It uses the [external data loader mechanism](https://www.rerun.io/docs/howto/open-any-file#external-dataloaders) to add this capability to the viewer without modifying the viewer itself.

https://github.com/rerun-io/rerun-loader-python-example-docx/assets/2624717/ceb324e5-ac64-49cc-b11d-eed5df5b7439

External data loaders are executables that are available to the Rerun Viewer via the `PATH` variable, with a name that starts with `rerun-loader-`.

This example is written in Python, and uses the [python-docx](https://python-docx.readthedocs.io/en/latest/) package to read `docx` files,
and some [ChatGPT](https://chat.openai.com/) crafted code to convert it to markdown format. 
The markdown formatted text is the logged to Rerun as a [`rr.TextDocument`](https://www.rerun.io/docs/reference/types/archetypes/text_document).

> ⚠️ NOTE: The actual docx -> markdown conversion code has only been tested on the `example.docx` file included in this repository and is only meant as a fun example.

## Installing the Rerun Viewer
The simplest option is just:
```bash
pip install rerun-sdk
```
Read [this guide](https://www.rerun.io/docs/getting-started/installing-viewer) for more options.

## Installing the plugin
### Installing pipx

The most robust way to install the plugin to your `PATH` is using [pipx](https://pipx.pypa.io/stable/).

If you don't have `pipx` installed on your system, you can follow the official instructions [here](https://pipx.pypa.io/stable/installation/).

### Installing the plugin with pipx
Now you can install the plugin to your `PATH` using

```bash
pipx install git@github.com:rerun-io/rerun-loader-python-example-docx.git
pipx ensurepath
```

Make sure it's installed by running it from your terminal, which should output an error and usage description:
```bash
rerun-loader-docx
usage: rerun-loader-docx [-h] [--recording-id RECORDING_ID] filepath
rerun-loader-docx: error: the following arguments are required: filepath
```

## Try it out
### Download an example docx file
```bash
curl -OL https://raw.githubusercontent.com/rerun-io/rerun-loader-python-example-docx/main/example.docx
```

### Open in the Rerun Viewer
You can either first open the viewer, and then open the file from there using drag-and-drop or the menu>open… dialog,
or you can open it directly from the terminal like:
```bash
rerun example.docx
```
