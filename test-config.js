// API設定テスト用の関数
function testApiConfiguration() {
  const config = {
    RAKUTEN_APP_ID: PropertiesService.getScriptProperties().getProperty('RAKUTEN_APP_ID'),
    YAHOO_APP_ID: PropertiesService.getScriptProperties().getProperty('YAHOO_APP_ID'),
    SCREENSHOTONE_ACCESS_KEY: PropertiesService.getScriptProperties().getProperty('SCREENSHOTONE_ACCESS_KEY'),
    DIFFBOT_TOKEN: PropertiesService.getScriptProperties().getProperty('DIFFBOT_TOKEN'),
    SCREENSHOT_FOLDER_ID: PropertiesService.getScriptProperties().getProperty('SCREENSHOT_FOLDER_ID')
  };
  
  console.log('=== API設定確認 ===');
  Object.entries(config).forEach(([key, value]) => {
    console.log(`${key}: ${value ? '設定済み' : '未設定'}`);
  });
  
  return config;
}

// 楽天API接続テスト
function testRakutenApi() {
  try {
    const appId = PropertiesService.getScriptProperties().getProperty('RAKUTEN_APP_ID');
    if (!appId) throw new Error('RAKUTEN_APP_ID未設定');
    
    const url = `https://app.rakuten.co.jp/services/api/IchibaItem/Search/20220601?applicationId=${appId}&keyword=テスト&hits=1`;
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    console.log('楽天API接続: 成功');
    console.log('取得件数:', data.Items?.length || 0);
    return true;
  } catch (e) {
    console.log('楽天API接続: エラー -', e.message);
    return false;
  }
}