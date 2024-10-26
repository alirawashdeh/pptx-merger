# Powerpoint Merger

Quickly and easily merge multiple Powerpoint files (PPTX only) into a single output file.

![Diagram](/screenshot.png)

# Installation

Ensure that .Net 8.0 is installed, then carry out the following steps:

```
git clone https://github.com/alirawashdeh/pptx-merger.git
cd pptx-merger
dotnet restore
```

# Usage

Place pptx files into the ```pptx-merger``` folder or any subfolder. There are two example files in the ```pptx-sanitiser/pptx``` folder.
Run the application using the following:

```
dotnet run
```

A message will be outputted for each file processed:

```
Copied file: dogs.pptx to: output.pptx 👌
Copied slides from: cats.pptx to: output.pptx 👌
```