figma.showUI(__html__, { width: 640, height: 360 });

const config = {
  templatesPage: "Templates",
  gap: 20,
  document: figma.currentPage.parent, // do not change
  // [name in figma, name in the spreadsheet]
  variantConfig: [
    ["car type", "car type"],
    ["alignment", "alignment"],
    ["version", "version"],
    ["disclaimer", "disclaimer"],
    ["call to action", "cta"],
    ["colour", "colour"]
  ],
  // [name in figma, name in the spreadsheet, type]
  fillConfig: [
    ["image",             "image",             "image"],
    ["iconisation-image", "iconisation-image", "image"],
    ["disclaimer",        "disclaimer_txt",    "text"],
    ["cta",               "cta_txt",           "text"],
    ["subline",           "subline_txt",       "text"],
    ["headline",          "headline_txt",      "text"]
  ],
  requiredFields: [
    'size',
    'reportingLabel',
    'car type',
    'alignment',
    'version',
    'disclaimer',
    'cta',
    'image',
    'colour'
  ]
}

figma.ui.onmessage = async msg => {
  // console.log(msg)
  if (msg.type === 'import-data') {
    try {
      const gap = config.gap

      const currentPage = figma.currentPage
      
      // predefined values to optimize loop
      const start =  msg.data[0]
      const { width, height } = getWidthAndHeight(start['size'])

      let prevWidth = width,
          prevHeight = height,
          x = 0 - prevWidth - gap,
          y = 0 - prevHeight - gap

      // load Fonts
      await figma.loadFontAsync({ family: "BMWTypeNext Pro", style: "Bold" })
      await figma.loadFontAsync({ family: "BMWTypeNext Pro", style: "Light" })
      await figma.loadFontAsync({ family: "BMWTypeNext Pro", style: "Regular" })

      // MAIN LOOP
      for (let i = 0; i < msg.data.length; i++) {
        const data = msg.data[i]

        if (!data['size']) continue // skip on empty rows

        const { width, height } = getWidthAndHeight(data['size'])

        ;[y, x] = getNextPos(prevWidth, prevHeight, width, height, gap, data['frame'], x, y)

        // get variant
        const variant = getVariant(data)

        const sizeWrapper = variant.createInstance()

        fillElementByData(sizeWrapper, data)

        // set x and y
        sizeWrapper.x = x
        sizeWrapper.y = y
        
        // set name
        sizeWrapper.name = data['reportingLabel']

        // append
        currentPage.appendChild(sizeWrapper)

        // set prev sizes
        prevWidth = width
        prevHeight = height
      }
      figma.closePlugin('Successfully imported')
    }
    catch (err) {
      console.log(err)
      figma.ui.postMessage({ type: 'json-check-error', message: err })
      // figma.closePlugin(err)
    }
  }
  else if (msg.type == 'update-data') {
    const currentPage = figma.currentPage
    const selection = currentPage.selection
    
    if (!selection.length) return figma.ui.postMessage({ type: 'no-selection' })

    try {
      // load Fonts
      await figma.loadFontAsync({ family: "BMWTypeNext Pro", style: "Bold" })
      await figma.loadFontAsync({ family: "BMWTypeNext Pro", style: "Light" })
      await figma.loadFontAsync({ family: "BMWTypeNext Pro", style: "Regular" })
  
      for (let index = 0; index < selection.length; index++) {
        const selected = selection[index]
        const data = getObjectByReportingLabel(msg.data, selected.name)
        
        if (!data) return figma.ui.postMessage({ type: 'no-element', name: selected.name })
        // get variant
        const variant = getVariant(data)

        const sizeWrapper = variant.createInstance()

        fillElementByData(sizeWrapper, data)

        // set x and y
        sizeWrapper.x = selected.x
        sizeWrapper.y = selected.y
        
        // set name
        sizeWrapper.name = selected.name

        // append
        currentPage.appendChild(sizeWrapper)
        selected.remove()
      }
      figma.closePlugin('Successfully updated')
    }
    catch (err) {
      console.log(err)
      figma.ui.postMessage({ type: 'json-check-error', message: err })
      // figma.closePlugin(err)
    }
  }
  else if (msg.type === 'json-check') {
    try {
      isJsonCorrect(msg.json)
      figma.ui.postMessage({ type: 'json-check-success', json: msg.json, function: msg.function })
    }
    catch (err) {
      figma.ui.postMessage({ type: 'json-check-error', message: err })
    }
  }
  else if (msg.type === 'local-storage-write') {
    config.document.setPluginData(msg.name, msg.image)
  }
  else if (msg.type === 'local-storage-read') {
    const data = msg.data

    for (let i = 0; i < data.length; i++) {
      if (!data[i].size) continue // skip empty rows

      for (const element of config.fillConfig) {
        if (element[2] == 'image') {
          const name = data[i][element[1]]
          if (!name) continue
          const image = config.document.getPluginData(name)
          if (!image) return figma.ui.postMessage({ type: 'read-error', name })
          data[i][element[1]] = image.split(',')[1] // split removes data:image from base64
        }
      }
    }

    figma.ui.postMessage({ type: 'read-success', data, function: msg.function })
  }
  else if (msg.type === 'local-storage-clear') {
    const keys = config.document.getPluginDataKeys()
    keys.forEach(key => {
      config.document.setPluginData(key, '')
    })
  }
  else if (msg.type === 'cancel') {
    figma.closePlugin()
  }
}

function getWidthAndHeight(size : string) {
  const data = size.match(/(\d{2,})[x|X](\d{2,})/)
  if (data == null) throw "Invalid size format in the spreadsheet."
  return {
    width: Number(data[1]),
    height: Number(data[2])
  }
}

function getChild(parent, name : string) {
  const node = parent.findOne(child => child.name == name)
  return node
}

function brToNewLine(br_text : string) {
  if (br_text == null) return ''
  return replaceAllRegex(br_text, '<(br|BR|Br|bR)>', '\n')
}

function escapeRegExp(string : string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
}

function replaceAll(str : string, find : string, replace : string) {
  return str.replace(new RegExp(escapeRegExp(find), 'g'), replace);
}

function replaceAllRegex(str : string, find : string, replace : string) {
  return str.replace(new RegExp(find, 'g'), replace);
}

// next line or move right
// if different size - next line
// if same size but frame is 1 - next line
// if same size and next frame - next frame
function getNextPos(prevWidth : number, prevHeight : number, width : number, height : number, gap : number, frame, x : number, y : number) {
  if (prevWidth != width && prevHeight != height) {
    // next line
    return [
      y += prevHeight + gap,
      x = 0
    ]
  }

  if (frame == 1 || frame == undefined || frame == '') {
    // next line
    return [
      y += prevHeight + gap,
      x = 0
    ]
  }
  
  // next frame
  return [
    y,
    x += width + gap
  ]
}

function isJsonCorrect(json : object[]) {
  const keys = config.requiredFields
  for (let i = 0; i < keys.length; i++) {
    for (let j = 0; j < json.length; j++) {
      const prop = json[j].hasOwnProperty(keys[i])
      if (!prop) throw `Required property "${keys[i]}" has not found in ${json[j]['reportingLabel']}.`

      if (isEmpty(json[j]['size'])) continue // skip empty rows

      const value = json[j][keys[i]]
      if (isEmpty(value)) throw `Required property "${keys[i]}" is empty in ${json[j]['reportingLabel']}.`
    }
  }
  return true

  function isEmpty(data) {
    if (data == null) return true
    if (data == undefined) return true
    if (data == '') return true
    return false
  }
}

function getObjectByReportingLabel(data: object[], reportingLabel : string) {
  for (let i = 0; i < data.length; i++) {
    if (data[i]['reportingLabel'] == reportingLabel) return data[i]
  }
  return null
}

function fillElementByData(object, data : object) {
  const parent = object.children[0]

  for (const element of config.fillConfig) {
    if (element[2] == "image") {
      // get image
      if (data[element[1]] == '' || data[element[1]] == ' ' || data[element[1]] == undefined ||data[element[1]] == null) continue
      
      const uint8array = figma.base64Decode(data[element[1]])
      const img = figma.createImage(uint8array)
      const hash = img.hash
      // fill image
      const imageFrame = getChild(parent, element[0])
      if (imageFrame != null) {
        imageFrame.fills = [{ type: 'IMAGE', scaleMode: 'FILL', imageHash: hash }]
      }
    } else if (element[2] == "text") {
      const textFrame = getChild(parent, element[0])
      // fill
      if (textFrame != null) {
        textFrame.characters = brToNewLine(data[element[1]])
      }
    } else {
      throw `Wrong data type specified in fillConfig for ${element[1]}.`
    }
  }
}

function getVariant(data : object) {
  const document = config.document
  const templatesPage = document.findChild(page => page.name == config.templatesPage)
  if (templatesPage == null) throw 'Can not find page with Templates. Please add it to your project.'
  
  // const templateSizeSet = templatesPage.findChild(child => child.name == data['size']) // does not work somewhy
  const templateSizeSet = getChild(templatesPage, data['size'])
  if (templateSizeSet == null) throw `Can not find template set for size ${data['size']}.`

  let variantQueue = ''
  for (const param of config.variantConfig) {
    variantQueue += `${param[0]}=${data[param[1]]}, `
  }
  variantQueue = variantQueue.slice(0, -2) // trim last 2 chars

  const variant = templateSizeSet.findChild(child => { return child.name == variantQueue })
  if (variant == null) throw `Can not find variant {${variantQueue}} for size ${data['size']}.`

  return variant
}
