const LANG_ATTR = ['display-name', 'short-description', 'long-description']

function __isContainTextNode (list) {
  var state = false
  for (var i = 0; i < list.length; i++) {
    if (list[i].type === 'text') {
      state = true
    }
  }
  return state
}

module.exports = {
  LANG_ATTR,
  __isContainTextNode
}