# Excel Art

This project takes an image as input and outputs an [Excel](https://en.wikipedia.org/wiki/Microsoft_Excel "Excel") spreadsheet, with each cell&apos;s background color being the RGB values from the image. In other words, it converts your images into Excel Art.

Idea from Matt Parker - [https://www.youtube.com/watch?v=UBX2QQHlQ_I](https://www.youtube.com/watch?v=UBX2QQHlQ_I)

## Getting Started

### Prerequisites

```
Git
Java
Maven
```

### Installing

Clone the GitHub repo.

``
git clone http://www.github.com/ChuckCaplan/ExcelArt
``

Compile the project with Maven.

```
cd ExcelArt
mvn package
```

## Run Excel Art
Running with Maven is easy, though you can also run with the java command directly.

``
mvn exec:java -Dexec.mainClass="com.chuckcaplan.excelart.Main" -Dexec.args="-i sample.jpg"
``

### Try Different Options
``
mvn exec:java -Dexec.mainClass="com.chuckcaplan.excelart.Main" -Dexec.args="--help"
``

 Usage: java com.chuckcaplan.excelart.Main [options]
  Options:
  \* -i
      The filesystem path or URL of the input image.
    -o
      The filename of the resulting .xlsx spreadsheet.
      Default: output.xlsx
    -h, --help
      Show this help.


## Notes

- It takes 66 seconds to process the sample JPG included with the application on a 2012 Macbook Air with 4GB RAM using the default JVM memory settings.

## Author

* **[Chuck Caplan](https://www.linkedin.com/in/charlescaplan/)**

See also the list of [contributors](https://github.com/your/project/contributors) who participated in this project.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
