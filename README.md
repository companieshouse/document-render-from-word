# word-template-service
Currently just a demonstration of substituting values into Word documents

## Requirements
In order to build word-template-service locally you will need the following:
- [Java 8](http://www.oracle.com/technetwork/java/javase/downloads/jdk8-downloads-2133151.html)
- [Maven](https://maven.apache.org/download.cgi)
- [Git](https://git-scm.com/downloads)

## Getting started
To build the project run

```mvn package```

## Running the demo
The template used by the demo is a copy of **certificate.doc** from CHIPS converted to docx format as **certificate.docx**

Run the main class **uk.gov.companieshouse.templateengine.word.WordDocumentProcessor**

java -cp target/word-template-service.jar uk.gov.companieshouse.templateengine.word.WordDocumentProcessor```

The rendered version of certificate will be placed in a sub-directory under **output**. Each run generates a new sub-directory which is named for a UUID generated when the code is run. 