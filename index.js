const fs = require('fs')
const convert = require('xml-js')
const Excel = require('exceljs')
const workbook = new Excel.Workbook()
const dataUtil = require('./util/data')
const htmlToText = require('html-to-text')
// const XLSX = require('xlsx')

/* extract */
const file = fs.readFileSync('./data.xml')
const result = convert.xml2json(file.toString(), { compact: false, spaces: 4 })
fs.writeFileSync('extract.json', result)
console.log('Processing...')

/* transform */
const { Product } = require('./model/master')
const json = JSON.parse(result)
const all_product = json.elements[0].elements.filter(x => x.name === 'product')
const model = all_product.map(item => {
  let product = new Product(item.attributes['product-id'])
  let data = {}
  item.elements.forEach(item => {
    if (item.elements) {
      // if contain type text then set text
      if (dataUtil.__isContainTextNode(item.elements)) {
        // if elements has lang attr ex: display-name
        if (dataUtil.LANG_ATTR.includes(item.name)) {
          data[item.name] = {
            lang: item.attributes['xml:lang'],
            text: item.elements[0].text
          }
        } else {
          data[item.name] = item.elements[0].text
        }
      } else {
        // if elements is images
        if (item.name === 'images') {
          // data[item.name] = item.elements.map(x => x.elements.filter(y => y.name === 'image').map(z => {
          //   return z.attributes.path
          // }))
          data[item.name] = item.elements.map(x => {
            const imageGroup = x.elements.map(y => {
              return {
                type: y.name,
                data: y.name === 'image' ? y.attributes.path : {
                  id: y.attributes['attribute-id'],
                  value: y.attributes.value
                }
              }
            })
            return {
              type: x.attributes['view-type'],
              data: imageGroup
            }
          })
        } else if (item.name === 'custom-attributes') {
          data[item.name] = item.elements.map(x => {
            return {
              id: x.attributes['attribute-id'],
              lang: x.attributes['attribute-id'] !== 'color' ? x.attributes['xml:lang'] : '',
              text: x.elements[0].text
            }
          })
        }
        else {
          data[item.name] = 'TBA'
        }
      }
    } else {
      data[item.name] = ''
    }
  })
  product.data = data
  return product
})

/* load */
workbook.xlsx.readFile('product_master.xlsx')
  .then(function() {
    let id = []
    let color = []
    var worksheet = workbook.getWorksheet('商品マスタ')

    // follow the same when mappling image
    var column = worksheet.getColumn('K').values

    // get color code column
    // both product_id and colorCode should match
    var codes = worksheet.getColumn('Z').values
    
    column.forEach((value, row) => {
      // search for item row based on product id
      const constraint = model.filter(x => String(value) === x.product_id.replace('_hk', ''))
      const mdRow = worksheet.getRow(row)

      let obj = {}
      obj.product_id = String(value)
      id.push(obj)

      // display name
      worksheet.getRow(6).getCell(16).value = 'display-name xml:lang="zh-HK"'
      worksheet.getRow(6).getCell(49).value = 'short-description xml:lang="zh-HK"'
      if (constraint.length > 0) {
        if (constraint[0].data['display-name'].lang === 'zh-HK') {
          // 16 is the column number of display-name
          mdRow.getCell(16).value = constraint[0].data['display-name'].text
        }

        if (constraint[0].data['short-description'].lang === 'zh-HK') {
          mdRow.getCell(49).value = htmlToText.fromString(constraint[0].data['short-description'].text, { wordwrap: 130 })
        }

        const ingredientFilter = constraint[0].data['custom-attributes'].filter(x => x.id === 'fullingredients')
        if (ingredientFilter.length > 0) {
          if (ingredientFilter[0].lang === 'zh-HK') {
            mdRow.getCell(58).value = htmlToText.fromString(ingredientFilter[0].text, { wordwrap: 130 })
          }
        }

        const subNameFilter = constraint[0].data['custom-attributes'].filter(x => x.id === 'subName')
        if (subNameFilter.length > 0) {
          if (subNameFilter[0].lang === 'zh-HK') {
            mdRow.getCell(19).value = subNameFilter[0].text
          }
        }

        const tab1TitleFilter = constraint[0].data['custom-attributes'].filter(x => x.id === 'tab1Title')
        if (tab1TitleFilter.length > 0) {
          if (tab1TitleFilter[0].lang === 'zh-HK') {
            mdRow.getCell(62).value = tab1TitleFilter[0].text
          }
        }

        const tab1ContentFilter = constraint[0].data['custom-attributes'].filter(x => x.id === 'tab1Content')
        if (tab1ContentFilter.length > 0) {
          if (tab1ContentFilter[0].lang === 'zh-HK') {
            mdRow.getCell(63).value = htmlToText.fromString(tab1ContentFilter[0].text, { wordwrap: 130 })
          }
        }

        const tab2TitleFilter = constraint[0].data['custom-attributes'].filter(x => x.id === 'tab2Title')
        if (tab2TitleFilter.length > 0) {
          if (tab2TitleFilter[0].lang === 'zh-HK') {
            mdRow.getCell(64).value = tab2TitleFilter[0].text
          }
        }

        const tab2ContentFilter = constraint[0].data['custom-attributes'].filter(x => x.id === 'tab2Content')
        if (tab2ContentFilter.length > 0) {
          if (tab2ContentFilter[0].lang === 'zh-HK') {
            mdRow.getCell(65).value = htmlToText.fromString(tab2ContentFilter[0].text, { wordwrap: 130 })
          }
        }

        const pdpMiddleContent1Filter = constraint[0].data['custom-attributes'].filter(x => x.id === 'pdpMiddleContent1')
        if (pdpMiddleContent1Filter.length > 0) {
          if (pdpMiddleContent1Filter[0].lang === 'zh-HK') {
            mdRow.getCell(66).value = htmlToText.fromString(pdpMiddleContent1Filter[0].text, { wordwrap: 130 })
          }
        }

        const pdpMiddleContent2Filter = constraint[0].data['custom-attributes'].filter(x => x.id === 'pdpMiddleContent2')
        if (pdpMiddleContent2Filter.length > 0) {
          if (pdpMiddleContent2Filter[0].lang === 'zh-HK') {
            mdRow.getCell(67).value = htmlToText.fromString(pdpMiddleContent2Filter[0].text, { wordwrap: 130 })
          }
        }
      } else {
        // ignore title
        if (row > 9) {
          mdRow.getCell(16).value = ''
          mdRow.getCell(49).value = ''

          // TODO: empty images
        }
      }
    })

    // color code workout
    codes.forEach((value, row) => {
      color.push(value)
    })
    id.forEach((item, i) => {
      id[i].colorCode = color[i]
    })
    
    // map images data
    id.forEach((value, row) => {
      const constraint = model.filter(x => String(value.product_id) === x.product_id.replace('_hk', ''))
      const mdRow = worksheet.getRow(row)
      if (constraint.length > 0) {
        const images = constraint[0].data.images
        images.forEach(image => {
          if (image.type === 'swatch') {
            // this is colorTip
            const variation = image.data.filter(x => x.type === 'variation')
            const imageFilter = image.data.filter(x => x.type === 'image')
            if (variation.length > 0 && imageFilter.length > 0) {
              if (value.colorCode === variation[0].data.value) {
                mdRow.getCell(30).value = imageFilter[0].data
              }
            }
          } else {
            // this is images data
            const variation = image.data.filter(x => x.type === 'variation')
            const imageFilter = image.data.filter(x => x.type === 'image')
            if (variation.length > 0 && imageFilter.length > 0) {
              if (value.colorCode === variation[0].data.value) {
                let image = ''
                imageFilter.forEach(data => {
                  image += `\n${data.data}`
                })
                mdRow.getCell(31).value = image
              }
            }
          }
        })
      }
    })
    return workbook.xlsx.writeFile('output/output.xlsx')
  })
  .catch(function(e) {
    console.log(e)
  })