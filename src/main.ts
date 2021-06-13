const sheets = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('sheet_id'))
const slackToken = PropertiesService.getScriptProperties().getProperty('slack_token')
const channelsSheetName = 'channels'
const usersSheetName = 'users'
const groupsSheetName = 'groups'

const saveSlackInfo = () => {
    saveChannels()
    saveUsers()
    saveGroups()
}

const fetchSlackInfo = (url: string, key: string, cursor: string, lasts: any[]): any[] => {
    const formData = {
        token: slackToken,
        cursor: cursor
    }
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        method: 'get',
        payload: formData,
        muteHttpExceptions: true
    }
    const responseJson = UrlFetchApp.fetch(url, options).getContentText()
    const response = JSON.parse(responseJson)
    const nextCursor = response.response_metadata?.next_cursor
    if (nextCursor && nextCursor.length > 0) {
        return fetchSlackInfo(url, key, nextCursor, lasts.concat(response[key]))
    } else {
        return lasts.concat(response[key])
    }
}

const saveChannels = () => {
    const channels = fetchSlackInfo('https://slack.com/api/conversations.list', 'channels', '', [])
    if (channels.length == 0) return
    const keys = Object.keys(channels[0])
    const values: any[] = channels.map(channel => {
        return keys.map(key => {
            if (['topic', 'purpose'].includes(key)) {
                return channel[key].value
            } else if (key == 'previous_names' && channel[key].length > 0) {
                return channel[key].join()
            } else {
                return channel[key]
            }
        })
    })
    values.unshift(keys)
    const sheet = sheets.getSheetByName(channelsSheetName)
    sheet.clear()
    sheet.getRange(1, 1, values.length, keys.length).setValues(values)
}

const saveUsers = () => {
    const users = fetchSlackInfo('https://slack.com/api/users.list', 'members', '', []).map(user => {
        const profileKeys = Object.keys(user.profile)
        profileKeys.forEach(key => user[key] = user.profile[key])
        delete user.profile
        return user
    })
    if (users.length == 0) return
    const keys = Object.keys(users[0])
    const values: any[] = users.map(channel => keys.map(key => channel[key]))
    values.unshift(keys)
    const sheet = sheets.getSheetByName(usersSheetName)
    sheet.clear()
    sheet.getRange(1, 1, values.length, keys.length).setValues(values)
}

const saveGroups = () => {
    const groups = fetchSlackInfo('https://slack.com/api/usergroups.list', 'usergroups', '', []).map(group => {
        const prefsKeys = Object.keys(group.prefs)
        prefsKeys.forEach(key => group[`prefs.${key}`] = group.prefs[key].join())
        delete group.prefs
        return group
    })
    if (groups.length == 0) return
    const keys = Object.keys(groups[0])
    const values: any[] = groups.map(channel => keys.map(key => channel[key]))
    values.unshift(keys)
    const sheet = sheets.getSheetByName(groupsSheetName)
    sheet.clear()
    sheet.getRange(1, 1, values.length, keys.length).setValues(values)
}

const setToken = () => {
    PropertiesService.getScriptProperties().setProperty('', '')
}