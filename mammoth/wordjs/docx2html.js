const mammoth = require('../');
const fs = require('fs');

const monospaceFonts = ["consolas", "courier", "courier new"];

let options = {
    transformDocument: mammoth.transforms.paragraph(transformParagraph),
    // preserveColors: true,
    // preserveFonts: true,
    styleMap: [
        
    ],
    // convertImage: mammoth.images.imgElement(function(image) {
    //     return image.read().then(function(buffer) {
    //         let file = new File([buffer], {type: image.contentType})
    //         let url = URL.createObjectURL(file);
    //         console.log(image);
    //         return {
    //             src: url,
    //             width: image.width,
    //             height: image.height
    //         };
    //     });
    // })
}
mammoth.convertToHtml({path: "word.docx"}, options)
.then(function(result){
    var html = result.value; // The generated HTML
    fs.writeFileSync('word.html', html);
})
.done();

function transformParagraph(paragraph) {
    var runs = mammoth.transforms.getDescendantsOfType(paragraph, "run");
    for(let run of runs){
        // console.log(run);
    }

    var isMatch = runs.length > 0 && runs.every(function(run) {
        return run.font && monospaceFonts.indexOf(run.font.toLowerCase()) !== -1;
    });
    if (isMatch) {
        return {
            ...paragraph,
            styleId: "code",
            styleName: "Code"
        };
    } else {
        return paragraph;
    }
}