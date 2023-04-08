const PDFServicesSdk = require('@adobe/pdfservices-node-sdk');
const fs = require('fs');
const path = require('path')
const XLSX = require('xlsx')
const { Configuration, OpenAIApi } = require('openai');
const readXlsxFile = require('read-excel-file/node');
require('dotenv').config();

// xlsx
// docx
// const OUTPUT = './output/noguti1_203.docx'

// sample1
// const INPUT = './input/noguti1_203.pdf'
// sample1
// const OUTPUT = 'output/noguti1_203.xlsx'

// sample2
// const INPUT = './input/yutaka_haits.pdf'
// sample2
// const OUTPUT = 'output/yutaka_haits.xlsx'

// sample3
// const INPUT = './input/kinshi_high_town.pdf'
// sample3
// const OUTPUT = 'output/kinshi_high_town.xlsx'

// sample4
const INPUT = './input/chitose_amflat.pdf'
// sample4
const OUTPUT = 'output/chitose_amflat.xlsx'
if(fs.existsSync(OUTPUT)) fs.unlinkSync(OUTPUT);

try {
    const credentials = PDFServicesSdk.Credentials.serviceAccountCredentialsBuilder().fromFile("PDFServicesSDK-Node.jsSamples/adobe-dc-pdf-services-sdk-node-samples/pdfservices-api-credentials.json").build();

    const executionContext = PDFServicesSdk.ExecutionContext.create(credentials);
    // This creates an instance of the Export operation we're using, as well as specifying output type (DOCX)
    const exportPdfOperation = PDFServicesSdk.ExportPDF.Operation.createNew(PDFServicesSdk.ExportPDF.SupportedTargetFormats.XLSX);

    // Set operation input from a source file
    const inputPDF = PDFServicesSdk.FileRef.createFromLocalFile(INPUT);
    exportPdfOperation.setInput(inputPDF);

    // execute
    exportPdfOperation.execute(executionContext)
    .then(async result => {
        console.log(result)
        await result.saveAsFile(OUTPUT)
    })
    .then(() => {
        let xlsxData = []
        readXlsxFile(fs.createReadStream(path.join(__dirname, OUTPUT))).then((rows) => {
            // `rows` is an array of rows
            // each row being an array of cells.
            rows = rows.map(d => {
                !!d
                d = d.filter(d => !!d)
                return d.flat()
            })
            xlsxData = rows
            //console.log(JSON.stringify(sheet))
            xlsxData = xlsxData.filter(data => data.length > 0)


            xlsxData = xlsxData.map(data => {
                if (!data) {
                    console.log("a")
                }
                let obj = {role: "user", content: data.join(",")}
                return obj
            })
            console.log(xlsxData)
            return new Promise((resolve) => {resolve(xlsxData)})
          })
          .then((data) => {
                const configuration = new Configuration({
                    apiKey: process.env.OPENAI_API_KEY
                    organization: process.env.OPENAI_ORG_ID
                })
                const openai = new OpenAIApi(configuration)
                openai.createChatCompletion({
                    model: 'gpt-3.5-turbo',
                    messages: [
                        {role: "user", content: "これから、ある物件の情報を渡します。渡す情報は構造は複数あるように見えますが実際は全部同じ物件の情報を表しています。あなたは、渡された情報からこの物件の住所と物件名称を抽出してください。"},
                        ...xlsxData
                    ]
                }).then(response => {
                    console.log(response.data.choices[0].message.content)
                    console.log(response.data.choices[0])
                })
            })
            .catch(err => {
                console.log('Exception encountered while executing operation', err);
            });
    })



    /*
    const options = new PDFServicesSdk.ExtractPDF.options.ExtractPdfOptions.Builder()
    .addElementsToExtract(PDFServicesSdk.ExtractPDF.options.ExtractElementType.TEXT, PDFServicesSdk.ExtractPDF.options.ExtractElementType.TABLES)
    .addElementsToExtractRenditions(PDFServicesSdk.ExtractPDF.options.ExtractRenditionsElementType.TABLES)
    .build();

    // Create a new operation instance.
    const extractPDFOperation = PDFServicesSdk.ExtractPDF.Operation.createNew(),
        input = PDFServicesSdk.FileRef.createFromLocalFile(
            'input/noguti1_203.pdf',
            PDFServicesSdk.ExtractPDF.SupportedSourceFormat.pdf
        );

    // Set operation input from a source file
    extractPDFOperation.setInput(input);

    // Set options
    extractPDFOperation.setOptions(options);

    extractPDFOperation.execute(executionContext)
        .then(result => result.saveAsFile('output/ExtractTextTableInfoWithTablesRenditionsFromPDF.zip'))
        .catch(err => {
            if(err instanceof PDFServicesSdk.Error.ServiceApiError
                || err instanceof PDFServicesSdk.Error.ServiceUsageError) {
                console.log('Exception encountered while executing operation', err);
            } else {
                console.log('Exception encountered while executing operation', err);
            }
        });
        */
} catch (err) {
    console.log('Exception encountered while executing operation', err);
}
