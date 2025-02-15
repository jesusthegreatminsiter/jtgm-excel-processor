
## How to Run Jar Locally

### API Created

```http
  GET /jtgm/excel/mgroup
```

#### 1. Download the latest jar file of the JTGM Excel Processor
https://github.com/ChinieBocalan/JTGM-Excel-Processor/releases/

#### 2. Open a terminal to the folder where the jar file has been saved and run the
```http
  java -jar JTGM-Excel-Processor-1.0-SNAPSHOT-jar-with-dependencies.jar
```

#### 3. Prepare a folder home directory
After create a folder in home directory `/JTGM MGroup/Raw` as it is stated the files to process. And add all the files to process.

#### 4. Run the following command in new terminal.
```http
  curl -XGET 'http://localhost:8080/jtgm/excel/mgroup'
```


## How to build executable jar

```
  mvn clean compile assembly:single
```


