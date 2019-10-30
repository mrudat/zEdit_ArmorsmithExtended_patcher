/* eslint-env node */
/* global xelib, registerPatcher, patcherPath, patcherUrl, info, fh */

const {
  AddArrayItem,
  AddElement,
  AddElementValue,
  AddKeyword,
  EditorID,
  ElementMatches,
  GetArrayItem,
  GetElementFile,
  GetElements,
  GetFileName,
  GetFormID,
  GetLinksTo,
  GetMasterRecord,
  GetRecordFlag,
  GetUIntValue,
  GetValue,
  HasElement,
  HasKeyword,
  IsWinningOverride,
  LongName,
  RemoveArrayItem,
  RemoveElement,
  RemoveKeyword,
  SetIntValue,
  SetUIntValue,
  SetValue
} = xelib

const {
  jetpack
} = fh

function pushOrAdd (obj, key, ...values) {
  if (obj[key] == null) {
    obj[key] = []
  }
  obj[key].push(...values)
}

function addToSet (obj, key, ...values) {
  if (obj[key] == null) {
    obj[key] = new Set()
  }
  for (const value of values) {
    obj[key].add(value)
  }
}

function getOrDefault (obj, key, value) {
  if (obj[key] == null) {
    obj[key] = value
  }
  return obj[key]
}

function mapGetOrDefault (map, key, value) {
  if (!map.has(key)) {
    map.set(key, value)
  }
  return map.get(key)
}

function AddAttachParent (record, keyword) {
  return AddArrayItem(record, 'APPR', '', keyword)
}

function RemoveAttachParent (record, keyword) {
  return RemoveArrayItem(record, 'APPR', '', keyword)
}

function csvToArray (text) {
  let previousCharacter = ''
  let field = ''
  let row = [field]
  const table = []
  let columnNumber = 0
  let rowNumber = 0
  let outsideQuote = true
  let character
  for (character of text) {
    if (character === '"') {
      if (outsideQuote && character === previousCharacter) {
        field += character
      }
      outsideQuote = !outsideQuote
    } else if (outsideQuote && character === ',') {
      row[columnNumber] = field
      field = ''
      columnNumber += 1
      character = ''
    } else if (outsideQuote && character === '\n') {
      if (previousCharacter === '\r') {
        field = field.slice(0, -1)
      }
      row[columnNumber] = field
      table[rowNumber] = row
      row = []
      field = ''
      rowNumber += 1
      columnNumber = 0
      character = ''
    } else {
      field += character
    }
    previousCharacter = character
  }
  if (field !== '') {
    row[columnNumber] = field
  }
  if (row.length) {
    table[rowNumber] = row
  }
  return table
}

function toCSV (data) {
  let file = ''

  let headings = new Set()
  for (const row of data) {
    for (const heading of Object.keys(row)) {
      headings.add(heading)
    }
  }
  headings = [...headings].sort().reverse()

  const headingsStart = headings.length - 1

  function maybeQuote (value) {
    if (value == null) return
    value = String(value)
    if (/[,"\n\r]/.test(value)) {
      file += '"' + value.replace('"', '""') + '"'
    } else {
      file += value
    }
  }

  function addRow (row) {
    let f
    if (row) {
      f = i => maybeQuote(row[headings[i]])
    } else {
      f = i => maybeQuote(headings[i])
    }
    for (let i = headingsStart; i >= 0; i--) {
      f(i)
      if (i > 0) file += ','
    }
    file += '\r\n'
  }

  addRow()
  for (const row of data) {
    addRow(row)
  }
  return file
}

async function loadCSV (path) {
  return csvToArray(await jetpack.cwd(patcherPath).readAsync(path))
}

function applyHeadings (data, extraHeadings) {
  const headings = data.shift()
  const result = []
  const template = {}
  for (const heading of headings) {
    Object.defineProperty(template, heading, {
      configurable: true,
      enumerable: true,
      writable: true
    })
  }
  if (extraHeadings) {
    for (const heading of extraHeadings) {
      Object.defineProperty(template, heading, {
        configurable: true,
        enumerable: true,
        writable: true
      })
    }
  }
  Object.seal(template)
  for (const row of data) {
    const obj = Object.create(template)
    for (let i = 0; i < row.length; i++) {
      obj[headings[i]] = row[i]
    }
    result.push(obj)
  }
  return result
}

function areBitsSet (value, mask) {
  return (value & mask) === mask
}

function slotListToMask (slotlist) {
  let mask = 0
  for (const slotNumber of slotlist.split(',')) {
    mask |= 1 << (Number.parseInt(slotNumber) - 30)
  }
  return mask
}

function maskToSlotList (mask) {
  const list = []
  for (var i = 30; i < 62; i++) {
    if (areBitsSet(mask, 1 << (i - 30))) {
      list.push(i)
    }
  }
  return list
}

function setToArray (set) {
  return Array.from(set.values())
}

function mapOn (array, key) {
  const dict = {}
  for (const row of array) {
    dict[row[key]] = row
  }
  return dict
}

function isVaultSuit (slotKeyword, classKeyword) {
  return slotKeyword === '_ClothesTypeUnderarmor_Slot33' && classKeyword === '_ClothingClassVault-Tec'
}

function isArmor (slotKeyword) {
  return slotKeyword.startsWith('_ArmorSlot')
}

function isHelmet (slotKeyword) {
  return isArmor(slotKeyword) && slotKeyword.endsWith('30')
}

function adjustCobjBasedOnArmorValue (armorValue, cobj, patchData) {
  // TODO what about energy armor?

  // TODO a discount for Leather + Steel used in construction?
  const ballisticFiberCount = ((armorValue * 3) / 10) >> 0

  const ingredient = GetArrayItem(cobj, 'FVPA - Components', 'Component', 'c_AntiBallisticFiber')

  if (ingredient === 0) {
    if (ballisticFiberCount !== 0) {
      patchData.ballisticFiberCount = ballisticFiberCount
    }
  } else {
    const ingredientCount = GetValue(ingredient, 'Count')
    if (ingredientCount !== ballisticFiberCount) {
      patchData.ballisticFiberCount = ballisticFiberCount
    }
  }

  const requiredPerkLevel = ((armorValue / 10) >> 0) + 1

  let totalPerkLevel = 0
  const perkLevels = {}

  if (HasElement(cobj, 'Conditions')) {
    const perks = GetElements(cobj, 'Conditions')
      .filter(c => ElementMatches(c, 'CTDA\\Function', 'HasPerk'))
      .map(c => EditorID(GetLinksTo(c, 'CTDA\\Parameter #1')))

    for (const perk of perks) {
      // TODO follow next-perk chain to determine level?
      const perkLevel = Number.parseInt(perk.substring(perk.length - 2))
      if (Number.isNaN(perkLevel)) continue
      const perkName = perk.substring(0, perk.length - 2)
      perkLevels[perkName] = {
        perk: perk,
        level: perkLevel
      }
      totalPerkLevel += perkLevel
    }
  }

  let missingPerkLevel = requiredPerkLevel - totalPerkLevel

  if (missingPerkLevel > 0) {
    for (const perkName of ['Armorer', 'Science']) {
      let level = 0
      if (perkLevels[perkName]) {
        let perk
        ({ perk, level } = perkLevels[perkName])
        if (level >= 4) continue
        addToSet(patchData, 'removePerk', perk)
      }
      while (missingPerkLevel > 0 && level < 4) {
        level++
        missingPerkLevel--
      }
      pushOrAdd(patchData, 'addPerk', `${perkName}${level.toString(10).padStart(2, '0')}`)
      if (missingPerkLevel === 0) break
    }
  }
}

function addSimpleObjectTemplate (armor) {
  if (!HasElement(armor, 'Object Template')) {
    AddElement(armor, 'Object Template')
  }

  if (!HasElement(armor, 'Object Template\\Combinations\\[0]')) {
    const combination = AddArrayItem(armor, 'Object Template\\Combinations', '', '')
    SetIntValue(combination, 'OBTS - Object Mod Template Item\\Addon Index', -1)
    SetValue(combination, 'OBTS - Object Mod Template Item\\Default', 'True')
  }
}

function isClassKeyword (keyword) {
  return /_(?:Armor|Clothing)Class/.test(keyword)
}

function isSlotKeyword (keyword) {
  return Object.prototype.hasOwnProperty.call(slotDataByKeyword, keyword)
}

function guessSlotKeyword (armor) {
  if (!HasElement(armor, 'BOD2')) {
    return '_ClothingSlotDevice'
  }

  const isArmored = GetValue(armor, 'FNAM\\Armor Rating') > 0

  const slotMask = GetUIntValue(armor, 'BOD2\\First Person Flags')

  const shortlist = slotDataList.filter(slotDatum => areBitsSet(slotMask, slotDatum.identifySlots))
  let slotDatum = shortlist.find(slotDatum => slotDatum.isArmored === isArmored)
  if (!slotDatum) {
    slotDatum = shortlist.shift()
  }
  if (slotDatum) {
    return slotDatum.keyword
  }

  return '_ClothingSlotDevice'
}

function wait (timeout) {
  return new Promise((resolve) => {
    setTimeout(() => {
      resolve()
    }, timeout)
  })
}

var afterXelibIsDone = (function () {
  const xelibIsDone = (async function () {
    console.log('Waiting for xelib to finish loading')
    while (xelib.GetLoaderStatus() <= 1) {
      await wait(1000)
    }
    console.log('xelib is done')
  })()
  return async function (func) {
    await xelibIsDone
    return func()
  }
})()

const PatchData = new Map()

const armorDataHeadings = [
  'armorEditorID',
  'slotKeyword',
  'classKeyword',
  'slotMask',
  'addsCarryWeight',
  'isHighTech'
]

async function loadPatchData () {
  const dataDir = jetpack.cwd(xelib.GetGlobal('DataPath')).dir('Tools').dir('ArmorsmithExtended_patcher')
  await Promise.all(dataDir.find('.', { matching: '*.csv' }).map(
    async function (filename) {
      const patchName = filename.substring(0, filename.length - 4)
      console.log(`Loading data for ${patchName} from ${filename}`)
      try {
        PatchData.set(patchName,
          mapOn(
            applyHeadings(
              csvToArray(await dataDir.readAsync(filename)),
              armorDataHeadings
            ),
            'armorEditorID'
          )
        )
      } catch (e) {
        console.log(`Couldn't load ${filename}: ${e.message}`)
      }
    }
  ))
}

afterXelibIsDone(loadPatchData)

let slotDataList

let slotDataByKeyword

async function loadSlotData () {
  const data = applyHeadings(await loadCSV('slotData.csv'))
  const slotKeys = [
    'identifySlots',
    'mandatorySlots',
    'allowedSlots'
  ]
  const boolKeys = [
    'isArmored',
    'isOutfit'
  ]
  for (const row of data) {
    for (const key of slotKeys) {
      if (row[key]) {
        row[key] = slotListToMask(row[key])
      } else {
        row[key] = undefined
      }
    }
    for (const key of boolKeys) {
      if (row[key]) {
        row[key] = row[key] === 'Y'
      } else {
        row[key] = undefined
      }
    }
  }
  slotDataList = data

  slotDataByKeyword = mapOn(slotDataList, 'keyword')
}

afterXelibIsDone(loadSlotData)

const forbiddenKeywords = new Set([
  'ma_armor_lining',
  'ma_VaultSuit',
  'ma_armor_Metal_Torso',
  'ma_armor_Lining_Leather_LimbArm',
  'ma_armor_Lining_Leather_LimbLeg'
])

const forbiddenAPs = new Set([
  'ap_armor_Lining'
])

const keywordTable = {
  38: {
    keywords: [
      'AEC_ma_armor_Addon',
      'AEC_ma_armor_Lining',
      'ma_Railroad_ClothingArmor'
    ],
    APs: [
      'AEC_ap_Addon',
      'AEC_ap_Lining',
      'ap_Railroad_ClothingArmor'
    ]
  },
  outfit: {
    keywords: [
      'AEC_ma_armor_Addon',
      'AEC_ma_armor_Lining',
      'ma_Railroad_ClothingArmor'
    ],
    APs: [
      'AEC_ap_Addon',
      'AEC_ap_Lining',
      'ap_Railroad_ClothingArmor'
    ]
  },
  47: {
    keywords: [
      'AEC_ma_armor_Eyewear'
    ],
    APs: [
      'AEC_ap_Eyewear'
    ]
  },
  34: {
    keywords: [
      'AEC_ma_armor_Glove'
    ],
    APs: [
      'AEC_ap_Glove'
    ]
  },
  30: {
    keywords: [
      'AEC_ma_armor_Lining',
      'AEC_ma_armor_Headgear_Addon',
      'ma_Railroad_ClothingArmor'
    ],
    APs: [
      'AEC_ap_AddonHeadgear',
      'AEC_ap_Lining',
      'ap_Railroad_ClothingArmor'
    ]
  },
  41: {
    keywords: [
      'ma_armor_lining'
    ],
    APs: [
      'ap_armor_Lining'
    ]
  },
  42: {
    keywords: [
      'ma_armor_lining'
    ],
    APs: [
      'ap_armor_Lining'
    ]
  },
  43: {
    keywords: [
      'ma_armor_lining'
    ],
    APs: [
      'ap_armor_Lining'
    ]
  },
  44: {
    keywords: [
      'ma_armor_lining'
    ],
    APs: [
      'ap_armor_Lining'
    ]
  },
  45: {
    keywords: [
      'ma_armor_lining'
    ],
    APs: [
      'ap_armor_Lining'
    ]
  }
}

const carryWeightKeywords = [
  'AEC_ma_armor_CarryWeight'
]

const carryWeightAPs = [
  'AEC_ap_CarryWeightBaseEffect',
  'AEC_ap_CarryWeightModifier'
]

const hasExtraCarryWeight = new Set([
  '_ClothingSlotBackpack_Slot54',
  '_ClothingSlotBandolier_Slot56',
  '_ClothingSlotBelt_Slot57',
  '_ClothingSlotPack_Slot54',
  '_ClothingSlotSatchel_Slot55',
  '_ClothingSlotTacticalVest_Slot57'
])

const globalAPs = [
  'ap_Legendary'
]

const blacklist = new Set([
  'AEC_One_Ring_To_Nude_Them_All',
  'AEC_One_Ring_To_Soil_Them_All',
  'AEC_One_Ring_For_The_Ghoulish'
])

registerPatcher({
  info: info,
  gameModes: [xelib.gmFO4],
  settings: {
    label: info.name,
    templateUrl: `${patcherUrl}/partials/settings.html`,
    defaultSettings: {
      patchFileName: 'zPatch.esp',
      ballisticWeaveOnlyForClothes: true
    }
  },
  requiredFiles: () => ['ArmorKeywords.esm', 'Armorsmith Extended.esp'],
  execute: (patchFile, helpers, settings, locals) => ({
    initialize: function () {
      locals.guesses = []

      locals.armas = new Map()
      locals.armors = new Map()

      locals.armoPatchData = new Map()
      locals.cobjPatchData = new Map()
    },
    process: [
      {
        load: {
          signature: 'ARMO',
          filter: function (armor) {
            if (!IsWinningOverride(armor)) return false
            if (GetRecordFlag(armor, 'Non-Playable')) return false
            if (!HasElement(armor, 'FULL')) return false
            if (!HasElement(armor, 'RNAM')) return false
            if (!ElementMatches(armor, 'RNAM', 'HumanRace')) return false
            if (HasElement(armor, 'KWDA')) {
              // TODO handle power armor?
              if (HasKeyword(armor, 'ArmorTypePower')) return false
            }

            const editorID = EditorID(armor)

            if (blacklist.has(editorID)) return false

            const patchData = Object.create(null)
            const addedKeywords = new Set()
            const removedKeywords = new Set()
            const addedAPs = new Set()
            const removedAPs = new Set()

            const fileName = GetFileName(GetElementFile(GetMasterRecord(armor)))

            const armorData = getOrDefault(mapGetOrDefault(PatchData, fileName, {}), editorID, {})

            if (armorData.name != null) {
              patchData.name = armorData.name
            }

            // TODO sanity check armorData keywords?

            let presentKeywords = new Set()

            if (HasElement(armor, 'KWDA')) {
              presentKeywords = new Set(GetElements(armor, 'KWDA').map(keyword => EditorID(GetLinksTo(keyword))))
            }

            let presentAPs = new Set()

            if (HasElement(armor, 'APPR')) {
              presentAPs = new Set(GetElements(armor, 'APPR').map(keyword => EditorID(GetLinksTo(keyword))))
            }

            const classKeywords = new Set()
            const slotKeywords = new Set()

            for (const keyword of presentKeywords) {
              if (forbiddenKeywords.has(keyword)) removedKeywords.add(keyword)
              if (isClassKeyword(keyword)) classKeywords.add(keyword)
              if (isSlotKeyword(keyword)) slotKeywords.add(keyword)
            }

            let slotKeyword = armorData.slotKeyword
            if (slotKeyword == null) slotKeyword = setToArray(slotKeywords).shift()
            if (slotKeyword == null) {
              slotKeyword = guessSlotKeyword(armor)
              locals.guesses.push(
                {
                  fileName: fileName,
                  armorEditorID: editorID,
                  slotKeyword: slotKeyword
                }
              )
              helpers.logMessage(`[WARN] Skipping ${LongName(armor)} as it has no slot keyword, but it could be ${slotKeyword}`)
              return false
            }

            let classKeyword = armorData.classKeyword
            if (classKeyword == null) classKeyword = setToArray(classKeywords).shift()

            const slotData = slotDataByKeyword[slotKeyword]

            if (slotData == null) {
              helpers.logMessage(`[ERROR] Skipping ${LongName(armor)} as we don't have data for ${slotKeyword}`)
              return false
            }

            const isAVaultSuit = isVaultSuit(slotKeyword, classKeyword)
            const isAnArmor = isArmor(slotKeyword)

            let currentSlotMask = 0
            let targetSlotMask

            if (HasElement(armor, 'BOD2')) {
              currentSlotMask = GetUIntValue(armor, 'BOD2\\First Person Flags')
            }

            if (armorData.slotMask != null) {
              targetSlotMask = slotListToMask(armorData.slotMask)
            } else {
              targetSlotMask = currentSlotMask
              targetSlotMask &= slotData.allowedSlots
              targetSlotMask |= slotData.mandatorySlots
            }

            if (targetSlotMask !== currentSlotMask) {
              patchData.slotMask = targetSlotMask
            }

            function applyKeyword (requiredKeyword, foundKeywords) {
              if (!foundKeywords.has(requiredKeyword)) {
                addedKeywords.add(requiredKeyword)
              } else {
                foundKeywords.delete(requiredKeyword)
                for (const keyword of foundKeywords) {
                  removedKeywords.add(keyword)
                }
              }
            }

            function ensureKeywords (keywords) {
              for (const keyword of keywords) {
                removedKeywords.delete(keyword)
                if (!presentKeywords.has(keyword)) addedKeywords.add(keyword)
              }
            }

            function ensureAPs (keywords) {
              for (const keyword of keywords) {
                removedAPs.delete(keyword)
                if (!presentAPs.has(keyword)) addedAPs.add(keyword)
              }
            }

            function removeKeywords (keywords) {
              for (const keyword of keywords) {
                addedKeywords.delete(keyword)
                if (presentKeywords.has(keyword)) removedKeywords.add(keyword)
              }
            }

            function removeAPs (keywords) {
              for (const keyword of keywords) {
                addedAPs.delete(keyword)
                if (presentAPs.has(keyword)) removedAPs.add(keyword)
              }
            }

            applyKeyword(slotKeyword, slotKeywords)

            if (classKeyword) {
              applyKeyword(classKeyword, classKeywords)
            }

            removeKeywords(forbiddenKeywords)
            removeAPs(forbiddenAPs)

            ensureAPs(globalAPs)

            const instanceNamingRules = EditorID(GetLinksTo(armor, 'INRD'))
            let targetInstanceNamingRules

            if (isAVaultSuit) {
              targetInstanceNamingRules = 'dn_VaultSuit'
            } else {
              if (isAnArmor) {
                if (isHelmet(slotKeyword)) {
                  // helmets use dn_Clothes INNR
                  targetInstanceNamingRules = 'dn_Clothes'
                } else {
                  targetInstanceNamingRules = 'dn_CommonArmor'
                }
              } else {
                targetInstanceNamingRules = 'dn_Clothes'
              }
            }

            if (instanceNamingRules !== targetInstanceNamingRules) {
              patchData.instanceNamingRules = targetInstanceNamingRules
            }

            const slots = new Set(maskToSlotList(targetSlotMask))

            if (slotData.isOutfit) {
              slots.add('outfit')
            }

            for (const slot of slots) {
              if (keywordTable[slot] == null) continue
              const { keywords, APs } = keywordTable[slot]
              ensureKeywords(keywords)
              ensureAPs(APs)
            }

            if ((armorData.addsCarryWeight === 'Y') || hasExtraCarryWeight.has(slotKeyword)) {
              ensureKeywords(carryWeightKeywords)
              ensureAPs(carryWeightAPs)
            } else {
              removeKeywords(carryWeightKeywords)
              removeAPs(carryWeightAPs)
            }

            if (armorData.isHighTech === 'Y' || classKeyword === '_ArmorClassThermOptics') {
              ensureKeywords(['AEC_ma_armor_ThermOptics'])
              ensureAPs(['AEC_ap_ThermOptics'])
            } else {
              removeKeywords(['AEC_ma_armor_ThermOptics'])
              removeAPs(['AEC_ap_ThermOptics'])
            }

            if (settings.ballisticWeaveOnlyForClothes) {
              if (isAnArmor) {
                removeKeywords(['ma_Railroad_ClothingArmor'])
                removeAPs(['ap_Railroad_ClothingArmor'])
              }
            }

            if (!(HasElement(armor, 'Object Template') && HasElement(armor, 'Object Template\\Combinations\\[0]'))) {
              patchData.addObjectTemplate = true
            } else {
              // TODO add available color swaps to object templates?
            }

            // record corrected slotMask against associated ARMA records.
            if (HasElement(armor, 'Models')) {
              GetElements(armor, 'Models').forEach(m => {
                // FIXME doesn't work?
                const arma = GetLinksTo(m, 'MODL')
                const armaData = mapGetOrDefault(locals.armas, GetFormID(arma), {
                  arma: arma,
                  slotMask: targetSlotMask
                })
                armaData.slotMask |= targetSlotMask
              })
            }

            if (addedKeywords.size) patchData.addedKeywords = addedKeywords
            if (removedKeywords.size) patchData.removedKeywords = removedKeywords
            if (addedAPs.size) patchData.addedAPs = addedAPs
            if (removedAPs.size) patchData.removedAPs = removedAPs

            if (Object.keys(patchData).length) {
              locals.armoPatchData.set(GetFormID(armor), patchData)
              return true
            }

            return false
          }
        },
        patch: function (armor) {
          helpers.logMessage(`Processing ${LongName(armor)}`)

          const {
            addedAPs,
            addedKeywords,
            addObjectTemplate,
            instanceNamingRules,
            name,
            removedAPs,
            removedKeywords,
            slotMask
          } = locals.armoPatchData.get(GetFormID(armor))

          if (name) {
            SetValue(armor, 'FULL', name)
          }

          if (removedKeywords && HasElement(armor, 'KWDA')) {
            for (const keyword of removedKeywords) {
              RemoveKeyword(armor, keyword)
            }
          }

          if (addedKeywords) {
            if (!HasElement(armor, 'KWDA')) AddElement(armor, 'KWDA')
            for (const keyword of addedKeywords) {
              AddKeyword(armor, keyword)
            }
          }

          if (removedAPs && HasElement(armor, 'APPR')) {
            for (const keyword of removedAPs) {
              RemoveAttachParent(armor, keyword)
            }
          }

          if (addedAPs) {
            if (!HasElement(armor, 'APPR')) AddElement(armor, 'APPR')
            for (const keyword of addedAPs) {
              AddAttachParent(armor, keyword)
            }
          }

          if (slotMask != null) {
            if (slotMask !== 0) {
              if (!HasElement(armor, 'BOD2')) {
                AddElement(armor, 'BOD2')
              }
              SetUIntValue(armor, 'BOD2\\First Person Flags', slotMask)
            } else {
              RemoveElement(armor, 'BOD2')
            }
          }

          if (instanceNamingRules) AddElementValue(armor, 'INRD', instanceNamingRules)

          if (addObjectTemplate) {
            try {
              addSimpleObjectTemplate(armor)
            } catch (e) {
              if (e.message.startsWith('Failed to add array item')) {
                helpers.logMessage('[WARN] Failed to add simple object template, instance naming rules may not work')
              } else {
                throw e
              }
            }
          }

          locals.armors.set(GetFormID(armor), {
            armor: armor
          })
        }
      },
      {
        records: function () {
          return Array.from(locals.armas.values())
            .filter(function (v) {
              const { arma, slotMask } = v
              let slots = 0
              if (HasElement(arma, 'BOD2')) slots = GetUIntValue(arma, 'BOD2\\First Person Flags')
              return (slots & slotMask) === 0
            })
            .mapOnKey('arma')
        },
        patch: function (arma) {
          helpers.logMessage(`Processing ${LongName(arma)}`)

          const { slotMask } = locals.armas.get(GetFormID(arma))

          if (slotMask !== 0) {
            if (!HasElement(arma, 'BOD2')) AddElement(arma, 'BOD2')
            SetUIntValue(arma, 'BOD2\\First Person Flags', slotMask)
          } else {
            // TODO have to have something?
            // RemoveElement(arma, 'BOD2')
          }
        }
      },
      {
        load: {
          signature: 'COBJ',
          filter: function (cobj) {
            if (!IsWinningOverride(cobj)) return false
            if (!HasElement(cobj, 'CNAM')) return false

            // TODO do all armor?
            const target = GetLinksTo(cobj, 'CNAM')

            if (target === 0) {
              helpers.logMessage(`Skipping ${LongName(cobj)} because it doesn't create anything! (CNAM is NULL)`)
              return false
            }

            const patchData = Object.create(null)

            const targetFormID = GetFormID(target)

            if (locals.armors.has(targetFormID)) {
              const { armor } = locals.armors.get(targetFormID)

              if (HasElement(armor, 'FNAM')) {
                const armorValue = GetValue(armor, 'FNAM\\Armor Rating')
                adjustCobjBasedOnArmorValue(armorValue, cobj, patchData)
              }

              if (!HasElement(cobj, 'INTV')) {
                patchData.addINTV = true
              }
            }

            // TODO COBJ for OMOD that adds armor?

            if (!Object.keys(patchData).length) return false

            locals.cobjPatchData.set(GetFormID(cobj), patchData)
            return true
          }
        },
        patch: function (cobj) {
          helpers.logMessage(`Processing ${LongName(cobj)}`)

          const {
            addINTV,
            ballisticFiberCount,
            addPerk,
            removePerk
          } = locals.cobjPatchData.get(GetFormID(cobj))

          if (addINTV) {
            AddElementValue(cobj, 'INTV\\Created Object Count', '1')
          }

          if (ballisticFiberCount != null) {
            let ingredient = GetArrayItem(cobj, 'FVPA - Components', 'Component', 'c_AntiBallisticFiber')
            if (ingredient === 0) {
              ingredient = AddArrayItem(cobj, 'FVPA - Components', 'Component', 'c_AntiBallisticFiber')
            }
            SetUIntValue(ingredient, 'Count', ballisticFiberCount)
          }

          if (removePerk) {
            GetElements(cobj, 'Conditions')
              .filter(c => ElementMatches(c, 'CTDA\\Function', 'HasPerk'))
              .filter(c => removePerk.has(EditorID(GetLinksTo(c, 'CTDA\\Parameter #1'))))
              .forEach(c => RemoveElement(c))
          }

          if (addPerk) {
            for (const perk of addPerk) {
              const condition = xelib.AddCondition(cobj, 'HasPerk', '10000000', '1', perk)
              SetValue(condition, 'CTDA\\Run On', 'Subject')
            }
          }
        }
      }
    ],
    finalize: function () {
      if (locals.guesses.length) {
        try {
          helpers.logMessage('Saving guesses...')
          const dataDir = jetpack.cwd(xelib.GetGlobal('DataPath')).dir('Tools').dir('ArmorsmithExtended_patcher')
          dataDir.write('guesses.csv', toCSV(locals.guesses))
        } catch (e) {
          helpers.logMessage(`[ERROR] Failed to save guesses: ${e.message}`)
        }
      }
    }
  })
})
