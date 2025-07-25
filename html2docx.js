import { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, ExternalHyperlink, 
    Table, TableRow, TableCell, WidthType, HeadingLevel, ConcreteNumbering, Numbering,
    HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, HorizontalPositionAlign, VerticalPositionAlign,
    TextWrappingType, TextWrappingSide } from 'docx';

import {Buffer} from 'buffer';

async function file_from_url(url, name, defaultType = 'image/jpeg'){
    try {
        console.log('file_from_url', url);
        const response = await fetch(url);
        const data = await response.blob();
        return new File([data], name, {
            type: data.type || defaultType,
        });
    } catch (error) {
        console.log(error);
        return new File([''], 'empty.txt', {
            type: 'text/plain'
        })
    }
}

async function array_buffer_from_url(url){
    let file = await file_from_url(url);
    return await file.arrayBuffer();
}

async function buffer_from_url(url){
    let array_buffer = await array_buffer_from_url(url);
    return Buffer.from(array_buffer);
}

async function get_intrinsic_image_size(url) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
    img.onerror = () => resolve({ width: 0, height: 0 });
    img.crossOrigin = 'anonymous';
    img.src = url;
  });
}

function scale_down(size, max_width=600){
    if(size.width <= max_width){
        return size;
    } else {
        let scale = max_width/size.width;
        return {width: size.width*scale, height: size.height*scale}
    }
}

function wrap_lines_in_p(html) {
  const lines = html
    .replaceAll('<br>', '\n')
    .replaceAll('<br/>', '\n')
    .replaceAll('<br />', '\n')
    .split('\n')
    .map(line => line.trim())
    .filter(line => line);

  return lines.map(line => `<p>${line}</p>`).join('\n');
}

async function canvas_to_image(canvas) {
  const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
  const url = URL.createObjectURL(blob);
  const img = new Image();
  await new Promise(resolve => {
    img.onload = resolve;
    img.src = url;
  });
  return img;
}

const COLORS = {
    "black": "000000",
    "silver": "c0c0c0",
    "gray": "808080",
    "white": "ffffff",
    "maroon": "800000",
    "red": "ff0000",
    "purple": "800080",
    "fuchsia": "ff00ff",
    "green": "008000",
    "lime": "00ff00",
    "olive": "808000",
    "yellow": "ffff00",
    "navy": "000080",
    "blue": "0000ff",
    "teal": "008080",
    "aqua": "00ffff",
    "aliceblue": "f0f8ff",
    "antiquewhite": "faebd7",
    "aquamarine": "7fffd4",
    "azure": "f0ffff",
    "beige": "f5f5dc",
    "bisque": "ffe4c4",
    "blanchedalmond": "ffebcd",
    "blueviolet": "8a2be2",
    "brown": "a52a2a",
    "burlywood": "deb887",
    "cadetblue": "5f9ea0",
    "chartreuse": "7fff00",
    "chocolate": "d2691e",
    "coral": "ff7f50",
    "cornflowerblue": "6495ed",
    "cornsilk": "fff8dc",
    "crimson": "dc143c",
    "cyan": "00ffff",
    "darkblue": "00008b",
    "darkcyan": "008b8b",
    "darkgoldenrod": "b8860b",
    "darkgray": "a9a9a9",
    "darkgreen": "006400",
    "darkgrey": "a9a9a9",
    "darkkhaki": "bdb76b",
    "darkmagenta": "8b008b",
    "darkolivegreen": "556b2f",
    "darkorange": "ff8c00",
    "darkorchid": "9932cc",
    "darkred": "8b0000",
    "darksalmon": "e9967a",
    "darkseagreen": "8fbc8f",
    "darkslateblue": "483d8b",
    "darkslategray": "2f4f4f",
    "darkslategrey": "2f4f4f",
    "darkturquoise": "00ced1",
    "darkviolet": "9400d3",
    "deeppink": "ff1493",
    "deepskyblue": "00bfff",
    "dimgray": "696969",
    "dimgrey": "696969",
    "dodgerblue": "1e90ff",
    "firebrick": "b22222",
    "floralwhite": "fffaf0",
    "forestgreen": "228b22",
    "gainsboro": "dcdcdc",
    "ghostwhite": "f8f8ff",
    "gold": "ffd700",
    "goldenrod": "daa520",
    "greenyellow": "adff2f",
    "grey": "808080",
    "honeydew": "f0fff0",
    "hotpink": "ff69b4",
    "indianred": "cd5c5c",
    "indigo": "4b0082",
    "ivory": "fffff0",
    "khaki": "f0e68c",
    "lavender": "e6e6fa",
    "lavenderblush": "fff0f5",
    "lawngreen": "7cfc00",
    "lemonchiffon": "fffacd",
    "lightblue": "add8e6",
    "lightcoral": "f08080",
    "lightcyan": "e0ffff",
    "lightgoldenrodyellow": "fafad2",
    "lightgray": "d3d3d3",
    "lightgreen": "90ee90",
    "lightgrey": "d3d3d3",
    "lightpink": "ffb6c1",
    "lightsalmon": "ffa07a",
    "lightseagreen": "20b2aa",
    "lightskyblue": "87cefa",
    "lightslategray": "778899",
    "lightslategrey": "778899",
    "lightsteelblue": "b0c4de",
    "lightyellow": "ffffe0",
    "limegreen": "32cd32",
    "linen": "faf0e6",
    "magenta": "ff00ff",
    "mediumaquamarine": "66cdaa",
    "mediumblue": "0000cd",
    "mediumorchid": "ba55d3",
    "mediumpurple": "9370db",
    "mediumseagreen": "3cb371",
    "mediumslateblue": "7b68ee",
    "mediumspringgreen": "00fa9a",
    "mediumturquoise": "48d1cc",
    "mediumvioletred": "c71585",
    "midnightblue": "191970",
    "mintcream": "f5fffa",
    "mistyrose": "ffe4e1",
    "moccasin": "ffe4b5",
    "navajowhite": "ffdead",
    "oldlace": "fdf5e6",
    "olivedrab": "6b8e23",
    "orange": "ffa500",
    "orangered": "ff4500",
    "orchid": "da70d6",
    "palegoldenrod": "eee8aa",
    "palegreen": "98fb98",
    "paleturquoise": "afeeee",
    "palevioletred": "db7093",
    "papayawhip": "ffefd5",
    "peachpuff": "ffdab9",
    "peru": "cd853f",
    "pink": "ffc0cb",
    "plum": "dda0dd",
    "powderblue": "b0e0e6",
    "rosybrown": "bc8f8f",
    "royalblue": "4169e1",
    "saddlebrown": "8b4513",
    "salmon": "fa8072",
    "sandybrown": "f4a460",
    "seagreen": "2e8b57",
    "seashell": "fff5ee",
    "sienna": "a0522d",
    "skyblue": "87ceeb",
    "slateblue": "6a5acd",
    "slategray": "708090",
    "slategrey": "708090",
    "snow": "fffafa",
    "springgreen": "00ff7f",
    "steelblue": "4682b4",
    "tan": "d2b48c",
    "thistle": "d8bfd8",
    "tomato": "ff6347",
    "turquoise": "40e0d0",
    "violet": "ee82ee",
    "wheat": "f5deb3",
    "whitesmoke": "f5f5f5",
    "yellowgreen": "9acd32"
}

export default async function html2docx(html, strict=false){

    let document = typeof html === 'string' ? (new DOMParser()).parseFromString(html,'text/html') : html;

    document.querySelectorAll('figure').forEach(figure => {
        const caption = figure.querySelector('figcaption');
        if (caption) {
            figure.after(caption);
        }
    });

    if(!strict){
        group_orphaned_elements(document.body);
    }

    console.log(document.body);

    let docx_elements = [];
    let nodes = Array.from(document.querySelectorAll('p,pre,table,h1,h2,h3,h4,h5,h6,ul,ol,div,figure,figcaption'));

    nodes = nodes.filter(node => {
        return !nodes
        .filter(el => el != node)
        .some(el => el.contains(node));
    });

    if(nodes.length == 0){
        document = (new DOMParser()).parseFromString(wrap_lines_in_p(document.body.innerHTML),'text/html');
        nodes = Array.from(document.querySelectorAll('p'));
    }

    for(let node of nodes){
        let instance = nodes.indexOf(node);

        if(['P', 'PRE', 'DIV', 'FIGURE', 'FIGCAPTION'].includes(node.nodeName)){
            docx_elements.push(await build_paragraph(node));

        } else if(node.nodeName == 'TABLE'){
            docx_elements.push(await build_table(node));
            
        } else if(['H1', 'H2', 'H3', 'H4', 'H5', 'H6'].includes(node.nodeName)){
            docx_elements.push(await build_heading(node));
        } else if(node.nodeName == 'UL'){
            docx_elements.push(...await build_ul(node, instance));
        } else if(node.nodeName == 'OL'){
            docx_elements.push(...await build_ol(node, instance));
        }
    }

    let docx = new Document({
        sections: [{
            children: docx_elements
        }],
        numbering:{
            config:[{
              reference: 'arabic',
              levels: [
                {
                    level: 0,
                    format: "decimal",
                    text: "%1",
                    alignment: AlignmentType.START,
                    style: {
                        paragraph: {
                            indent: { left: 300, hanging: 200 },
                        },
                    },
                },
                {
                    level: 1,
                    format: "decimal",
                    text: "%1.%2",
                    alignment: AlignmentType.START,
                    style: {
                        paragraph: {
                            indent: { left: 600, hanging: 200 },
                        },
                    },
                },
              ],
            }]
          },
    })

    let blob = await Packer.toBlob(docx);
    return blob;
}


async function build_paragraph(node){
    let style = parse_style(node);
    if(style.size == null) style.size = 24;
    if(style.font == null && node.nodeName == 'PRE') style.font = 'Courier New';
    
    let children = await build_child_nodes(node, style);
    
    let alignment = get_align(node);
    let border = parse_border(node);
    if(node.parentElement.nodeName == 'BLOCKQUOTE'){
        if(border.left == null){
            border.left = {color: 'cbd5e1', size: 16, space: 1, style: 'single'}
        }
        if(style.indent == null || style.indent.left == 0){
            style.indent = {left: 80}
        }
    }

    let paragraph = new Paragraph({
        alignment,
        indent: style.indent,
        children,
        border
    });
    return paragraph;
}
async function build_table(node){
    let rows = [];
    for(let row of node.querySelectorAll('tr')){
        let cells = [];
        for(let cell of row.querySelectorAll('th, td')){
            
            cells.push(new TableCell({
                children: [new Paragraph({children: await build_child_nodes(cell)})]
            }));
        }
        rows.push(new TableRow({
            children: cells
        }))
    }
    let number_of_columns = node.querySelector('tr').querySelectorAll('th, td').length;
    let table = new Table({
        rows,
        width: 0,
        columnWidths: Array(number_of_columns).fill(Math.floor(9638/number_of_columns), 0, number_of_columns)
    });
    return table;
}

async function build_heading(node){
    let style = parse_style(node);
    let children = await build_child_nodes(node, style);
    let alignment = get_align(node);
    let heading;
    switch (node.nodeName) {
        case 'H1':
            heading = HeadingLevel.HEADING_1
            break;
        case 'H2':
            heading = HeadingLevel.HEADING_2;
            break;
        case 'H3':
            heading = HeadingLevel.HEADING_3;
            break;
        case 'H4':
            heading = HeadingLevel.HEADING_4;
            break;
        case 'H5':
            heading = HeadingLevel.HEADING_5;
            break;
        case 'H6':
            heading = HeadingLevel.HEADING_6;
            break;
        default:
            break;
    }
    let border = parse_border(node);
    let paragraph = new Paragraph({
        alignment,
        children,
        indent: style.indent,
        heading,
        border
    });
    return paragraph;
}

async function build_ul(node, instance){
    let list = [];
    for(let li of node.querySelectorAll(':scope > li')){
        list.push(new Paragraph({
            children: await build_child_nodes(li),
            bullet: {level: 0, instance}
        }))
        if(li.querySelectorAll == null) continue;
        for(let sub_li of li.querySelectorAll('ul li')){
            list.push(new Paragraph({
                children: await build_child_nodes(sub_li),
                bullet: {level: 1, instance}
            }))
        }
    }
    return list;
}

async function build_ol(node, instance){
    let list = [];
    for(let li of node.querySelectorAll(':scope > li')){
        list.push(new Paragraph({
            children: await build_child_nodes(li),
            numbering: {reference: 'arabic', instance, level: 0}
        }))
        if(li.querySelectorAll == null) continue;
        for(let sub_li of li.querySelectorAll('ol li')){
            list.push(new Paragraph({
                children: await build_child_nodes(sub_li),
                numbering: {reference: 'arabic', instance, level: 1}
            }))
        }
    }
    return list;
}

async function build_child_nodes(node, inherit_attr){

    if(inherit_attr == null) inherit_attr = {};
    let values = [];
    let children = node.childNodes;
    for(let child of children){
        if(child.nodeName == '#text'){
            let text_run = new TextRun({
                text: child.nodeValue,
                bold: inherit_attr.bold,
                italics: inherit_attr.italics,
                subScript: inherit_attr.subScript,
                superScript: inherit_attr.superScript,
                strike: inherit_attr.strike,
                underline: inherit_attr.underline ? {} : null,
                color: inherit_attr.color,
                shading: inherit_attr.shading,
                size: inherit_attr.size,
                allCaps: inherit_attr.allCaps,
                font: inherit_attr.font,
                style: inherit_attr.style
            })
            values = [...values, text_run];
        } else if(child.nodeName == 'A' && child.getAttribute('href')){
            let link = new ExternalHyperlink({
                children: await build_child_nodes(child, {style: 'Hyperlink'}),
                link: child.getAttribute('href')
            })
            values = [...values, link];
        } else if(child.nodeName == 'IMG') {
            let buffer = await buffer_from_url(child.src);
            let floating = parse_image_floating(child);
            
            let size = {width: child.width, height: child.height};
            let intrinsic_size = await get_intrinsic_image_size(child.src);

            if(size.width == 0 && size.height == 0){
               size = intrinsic_size;
               
            } else if(size.width == 0){
                let factor = intrinsic_size.width/intrinsic_size.height;
                size.width = size.height*factor;

            }  else if(size.height == 0){
                let factor = intrinsic_size.height/intrinsic_size.width;
                size.height = size.width*factor;
            }

            let image_run = new ImageRun({
                data: buffer,
                transformation: scale_down(size),
                floating
            })
            
            values = [...values, image_run];
        } else if(child.nodeName == 'CANVAS') {

            let image = await canvas_to_image(child);
            let buffer = await buffer_from_url(image.src);
            let floating = parse_image_floating(image);
            
            let size = {width: image.width, height: image.height};
            let intrinsic_size = await get_intrinsic_image_size(image.src);

            if(size.width == 0 || size.width == null){
               size = intrinsic_size;
            }

            let image_run = new ImageRun({
                data: buffer,
                transformation: scale_down(size),
                floating
            })
            
            values = [...values, image_run];
        } else if(node.childNodes.length > 0 && !['UL', 'OL'].includes(node.nodeName)){
            let passed_down_style = {...inherit_attr, ...parse_style(child)};

            if(child.nodeName == 'STRONG') passed_down_style.bold = true;
            if(child.nodeName == 'EM') passed_down_style.italics = true;
            if(child.nodeName == 'SUB') passed_down_style.subScript = true;
            if(child.nodeName == 'SUP') passed_down_style.superScript = true;
            if(child.nodeName == 'S') passed_down_style.strike = true;
            if(child.nodeName == 'U') passed_down_style.underline = true;

            if(child.nodeName == 'A') {
                passed_down_style.anchor = child.getAttribute('href');
                passed_down_style.underline = true;
            }

            values = [...values, ...await build_child_nodes(child, passed_down_style)];
        }
    }
    return values;
}

function parse_style(node){
    
    let style = {};

    let raw_style = (node.getAttribute('style') || '').split(';');

    for(let el of raw_style){
        let values = el.trim().split(':');
        if(values.length == 2){
            style[values[0].trim()] = values[1].trim();
        }
    }

    let fill = to_hex(style['background-color']);
    if(fill){
        style['shading'] = {fill};
    }

    style['color'] = to_hex(style['color']);
    
    if(style['font-family']){
        style['font'] = style['font-family'].split(',')[0];
        if(style['font']){
            style['font'] = style['font'].split('\'').join('');
        }
    }

    style['size'] = to_halfpoint(style['font-size']);
    let indent_left = to_halfpoint(style['padding-left']);
    if(!isNaN(indent_left)){
        //indent_left is in halfpoint, 
        //indentations in OpenXML are measured in 1/20 of a point
        style['indent'] = {left: indent_left*10}
    }

    if(style['text-transform'] == 'uppercase'){
        style['allCaps'] = true;
    }
    if(style['text-transform'] == 'capitalize'){
        style['smallCaps'] = true;
    }
    if(style['text-decoration'] == 'line-through'){
        style['strike'] = true;
    }
    if(style['text-decoration'] == 'underline'){
        style['underline'] = true;
    }
    if(style['font-style'] == 'italic'){
        style['italics'] = true;
    }
    if(style['font-weight'] == 'bold' || parseInt(style['font-weight']) >= 700){
        style['bold'] = true;
    }

    
    let allow_attrs = ['color', 'shading', 'size', 'indent', 'allCaps', 'allCaps', 'strike', 'font', 'italics','underline', 'bold'];
    for(let key of Object.keys(style)){
        if(!allow_attrs.includes(key) || style[key] == null){
            delete style[key];
        }
    }
    return style;
}

function to_halfpoint(str){
    if(str == null || str == '') return null;
    str = str.trim();

    let unit;
    if(str.endsWith('pt')){
        unit = 'pt'
    } else if(str.endsWith('px')){
        unit = 'px';
    }
    if(unit){
        let value = parseInt(str.split(unit).join(''));
        if(isNaN(value)) return null;
        if(unit == 'px'){
            value = 2*Math.ceil((72*value)/96)
        } else if(unit == 'pt'){
            value = 2*value;
        }
        return value;
    } else {
        return null;
    }
}

function to_hex(str){
    if(str == null || str == '') return null;
    str = str.trim();

    if(COLORS[str] != null) return COLORS[str];

    let color;
    if(str.includes('rgb')){
        color = rgb_to_hex(str);
    } else {
        color = str.split('#').join('').trim();
    }
    
    if(color.length == 3){
        color = color + color;
    }
    if(color == null || !/^[0-9A-F]{6}$/i.test(color)){
        color == null
    }
    return color;
}
function rgb_to_hex(rgb) {
    // Choose correct separator
    let sep = rgb.indexOf(",") > -1 ? "," : " ";
    // Turn "rgb(r,g,b)" into [r,g,b]
    rgb = rgb.substr(4).split(")")[0].split(sep);

    let r = (+rgb[0]).toString(16),
        g = (+rgb[1]).toString(16),
        b = (+rgb[2]).toString(16);

    if (r.length == 1)
        r = "0" + r;
    if (g.length == 1)
        g = "0" + g;
    if (b.length == 1)
        b = "0" + b;

    return r + g + b;
}

function parse_border(node){
    let style = {};
    let raw_style = (node.getAttribute('style') || '').split(';');

    for(let el of raw_style){
        let values = el.trim().split(':');
        if(values.length == 2){
            style[values[0].trim()] = values[1].trim();
        }
    }
    let top, right, bottom, left;
    if(style['border'] != null){
        let [size, border_style, ...color] = style['border'].split(' ');
        color = to_hex(color.join('').trim());
        size = to_halfpoint(size)*4;
        top = {color,size, space:1,style: 'single'}
        right = {color,size, space:1,style: 'single'}
        bottom = {color,size, space:1,style: 'single'}
        left = {color,size, space:1,style: 'single'}
    }
    if(style['border-left'] != null){
        let [size, border_style, ...color] = style['border-left'].split(' ');
        color = to_hex(color.join('').trim());
        size = to_halfpoint(size)*4;
        left = {color,size, space:1,style: 'single'}
    }
    if(style['border-right'] != null){
        let [size, border_style, ...color] = style['border-right'].split(' ');
        color = to_hex(color.join('').trim());
        size = to_halfpoint(size)*4;
        right = {color,size, space:1,style: 'single'}
    }
    if(style['border-top'] != null){
        let [size, border_style, ...color] = style['border-top'].split(' ');
        color = to_hex(color.join('').trim());
        size = to_halfpoint(size)*4;
        top = {color,size, space:1,style: 'single'}
    }
    if(style['border-bottom'] != null){
        let [size, border_style, ...color] = style['border-bottom'].split(' ');
        color = to_hex(color.join('').trim());
        size = to_halfpoint(size)*4;
        bottom = {color,size, space:1,style: 'single'}
    }

    return {top, right, bottom, left};
}

function parse_image_floating(node){
    let style = {};

    let raw_style = (node.getAttribute('style') || '').split(';');

    for(let el of raw_style){
        let values = el.trim().split(':');
        if(values.length == 2){
            style[values[0].trim()] = values[1].trim();
        }
    }
    let verticalPosition =  {
        relative: VerticalPositionRelativeFrom.TOP_MARGIN,
        align: VerticalPositionAlign.TOP
    }
    
    let margin = {
        top: 360000,
        right: 360000,
        bottom: 360000,
        left: 360000
    }
    if(style['margin-left'] == 'auto' && style['margin-right'] == 'auto'){
        return {
            horizontalPosition: {
                relative: HorizontalPositionRelativeFrom.COLUMN,
                align: HorizontalPositionAlign.CENTER,
            },
            verticalPosition, margin,
            wrap: {type: TextWrappingType.TOP_AND_BOTTOM, side: TextWrappingSide.BOTH_SIDES}
        }
    }
    if(style['float'] == 'left'){
        return {
            horizontalPosition: {
                relative: HorizontalPositionRelativeFrom.COLUMN,
                align: HorizontalPositionAlign.LEFT,
            },
            verticalPosition, margin,
            wrap: {type:TextWrappingType.SQUARE, side: TextWrappingSide.RIGHT}
        }
    }
    if(style['float'] == 'right'){
        return {
            horizontalPosition: {
                relative: HorizontalPositionRelativeFrom.COLUMN,
                align: HorizontalPositionAlign.RIGHT,
            },
            verticalPosition, margin,
            wrap: {type:TextWrappingType.SQUARE, side: TextWrappingSide.LEFT}
        }
    }
    return null;
}



function get_align(node){
    let raw_style = node.getAttribute('style');
    if(raw_style == null) return AlignmentType.LEFT;
    let style = {};
    for(let pair of raw_style.split(';')){
        if(pair.split(':').length != 2) continue;
        let key = pair.split(':')[0].trim();
        let value = pair.split(':')[1].trim();
        style[key] = value;
    }
    switch (style['text-align']) {
        case 'left':
            return AlignmentType.LEFT;
        case 'center':
            return AlignmentType.CENTER;
        case 'justify':
            return AlignmentType.JUSTIFIED;
        case 'right':
            return AlignmentType.RIGHT;
        default:
            return AlignmentType.LEFT;
    }
}



function group_orphaned_elements(container) {
    // Define block-level elements that are considered valid parents
    const validParents = new Set(['p', 'pre', 'div', 'table', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'figure', 'figcaption']);

    // Check if an element is truly orphaned (doesn't have valid parent/ancestor)
    function isOrphaned(element) {
        let current = element.parentElement;
        while (current && current !== container) {
            if (validParents.has(current.tagName.toLowerCase())) {
                return false;
            }
            current = current.parentElement;
        }
        return true;
    }

    // Check if an element is inline or inline-block
    function isInlineElement(element) {
        if (element.nodeType === Node.TEXT_NODE) {
            return true;
        }
        if (element.nodeType !== Node.ELEMENT_NODE) {
            return false;
        }

        const tagName = element.tagName.toLowerCase();
        const inlineElements = new Set([
            'a', 'abbr', 'acronym', 'b', 'bdo', 'big', 'br', 'button', 'cite', 'code',
            'dfn', 'em', 'i', 'img', 'input', 'kbd', 'label', 'map', 'object', 'q',
            'samp', 'script', 'select', 'small', 'span', 'strong', 'sub', 'sup',
            'textarea', 'time', 'tt', 'var'
        ]);

        return inlineElements.has(tagName) ||
            window.getComputedStyle(element).display.includes('inline');
    }

    // Check if a text node contains only whitespace
    function isWhitespaceOnly(textNode) {
        return /^\s*$/.test(textNode.textContent);
    }

    // Get all direct children of the container
    function processChildren(parent) {
        const children = Array.from(parent.childNodes);
        let i = 0;

        while (i < children.length) {
            const child = children[i];

            // Skip if this child is no longer in the DOM (might have been moved)
            if (!parent.contains(child)) {
                i++;
                continue;
            }

            // If it's a block element, recursively process its children
            if (child.nodeType === Node.ELEMENT_NODE && !isInlineElement(child)) {
                processChildren(child);
                i++;
                continue;
            }

            // Check if this is an orphaned inline element or text node
            const isChildOrphaned = isOrphaned(child);
            const isChildInline = isInlineElement(child);
            const isSignificantText = child.nodeType === Node.TEXT_NODE && !isWhitespaceOnly(child);

            if (isChildOrphaned && (isChildInline || isSignificantText)) {
                // Found start of orphaned inline sequence
                const orphanedGroup = [];
                let j = i;

                // Collect all consecutive orphaned inline elements
                while (j < children.length) {
                    const currentChild = children[j];

                    // Skip if this child is no longer in the DOM
                    if (!parent.contains(currentChild)) {
                        j++;
                        continue;
                    }

                    const isCurrentOrphaned = isOrphaned(currentChild);
                    const isCurrentInline = isInlineElement(currentChild);
                    const isCurrentSignificantText = currentChild.nodeType === Node.TEXT_NODE && !isWhitespaceOnly(currentChild);

                    // Special handling for <br> elements - they break the group
                    if (currentChild.nodeType === Node.ELEMENT_NODE &&
                        currentChild.tagName.toLowerCase() === 'br' &&
                        isCurrentOrphaned) {

                        // Add the <br> to current group
                        orphanedGroup.push(currentChild);
                        j++;

                        // Break the group here - we'll create a paragraph and start a new group
                        break;
                    }

                    // Special handling for orphaned <img> elements
                    if (currentChild.nodeType === Node.ELEMENT_NODE &&
                        currentChild.tagName.toLowerCase() === 'img' &&
                        isCurrentOrphaned) {

                        const imgDisplay = window.getComputedStyle(currentChild).display;

                        // If img is block-level, treat it separately
                        if (imgDisplay === 'block' || imgDisplay === 'block-inline') {
                            // If we have content in current group, close it first
                            if (orphanedGroup.length > 0 && orphanedGroup.some(node =>
                                node.nodeType === Node.ELEMENT_NODE ||
                                (node.nodeType === Node.TEXT_NODE && !isWhitespaceOnly(node)))) {
                                break; // This will close current group, then img will be handled separately
                            }

                            // Create separate paragraph for block img
                            const imgP = document.createElement('p');
                            parent.insertBefore(imgP, currentChild);
                            imgP.appendChild(currentChild);

                            // Update arrays and continue
                            const newChildren = Array.from(parent.childNodes);
                            const imgPIndex = newChildren.indexOf(imgP);
                            i = imgPIndex + 1;
                            children.length = 0;
                            children.push(...newChildren);
                            break;
                        } else {
                            // Inline img - treat like other inline elements
                            orphanedGroup.push(currentChild);
                            j++;
                        }
                    }
                    // If it's orphaned and inline, add to group
                    else if (isCurrentOrphaned && (isCurrentInline || isCurrentSignificantText)) {
                        orphanedGroup.push(currentChild);
                        j++;
                    }
                    // If it's just whitespace, include it but don't break the sequence
                    else if (currentChild.nodeType === Node.TEXT_NODE && isWhitespaceOnly(currentChild)) {
                        orphanedGroup.push(currentChild);
                        j++;
                    }
                    // If we hit a block element or non-orphaned element, stop
                    else {
                        break;
                    }
                }

                // Only wrap if we have actual content (not just whitespace)
                const hasSignificantContent = orphanedGroup.some(node =>
                    node.nodeType === Node.ELEMENT_NODE ||
                    (node.nodeType === Node.TEXT_NODE && !isWhitespaceOnly(node))
                );

                if (hasSignificantContent && orphanedGroup.length > 0) {
                    // Create a new paragraph element
                    const p = document.createElement('p');

                    // Insert the paragraph before the first orphaned element
                    parent.insertBefore(p, orphanedGroup[0]);

                    // Move all orphaned elements into the paragraph
                    orphanedGroup.forEach(node => {
                        if (parent.contains(node)) {
                            p.appendChild(node);
                        }
                    });

                    // Update the children array since DOM has changed
                    const newChildren = Array.from(parent.childNodes);
                    const pIndex = newChildren.indexOf(p);
                    i = pIndex + 1;
                    children.length = 0;
                    children.push(...newChildren);

                    // If the last element was a <br>, continue collecting for next group
                    if (j < children.length) {
                        continue;
                    }
                } else {
                    i = j;
                }
            } else {
                i++;
            }
        }
    }

    // Start processing from the container
    processChildren(container);
}
