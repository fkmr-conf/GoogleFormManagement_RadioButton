
const ssID = "1AqEjTm_4EaqiYFTcBJyakpTpYfFTLfDfrBnxYYHoXLE";
const ss = SpreadsheetApp.openById(ssID);
const response_sheet = ss.getSheetByName("フォームの回答 1");
const capacity_sheet = ss.getSheetByName("各日程の定員管理") ;

//const formID = "1Hlq_noumAAFOmMyB4aVtBmIfCsm2x58TbiYRDrIP63o";//URLをメール文で使用するため、今回はIDでは取得しない
const formURL = "https://docs.google.com/forms/d/1Hlq_noumAAFOmMyB4aVtBmIfCsm2x58TbiYRDrIP63o/edit";
const form = FormApp.openByUrl(formURL);
const form_title = form.getTitle;


function ManageOptions() {
  //フォーム回答送信のタイミングで実行

  //@「回答」シート
  //1.重複回答の処理、応募数のカウント
  let counter_obj = CountApplicationNumber();
  
  //@「各日程の定員管理」シート
  //2.応募数を入力しつつ、受付可能な選択肢のリストを取得
  let available_options = UpdateCapacityAndGetOptionList(counter_obj);

  //@フォーム
  //3.formの選択肢(2つめの質問)に反映
  let questions = form.getItems(); //対象フォームの全質問取得
  let target_question = questions[1]; //2番目の質問を取得※0からのカウントで何番目か
  target_question.asMultipleChoiceItem().setChoiceValues(available_options); //選択肢をセット

  //4.formの説明を作成
  let description_text = CreateDescription();

  //5.formの説明を更新
  form.setDescription(description_text);

}


//1.重複回答の処理、応募数のカウント
function CountApplicationNumber(){
  let last_row = response_sheet.getLastRow(); //最終行数
  let latest_respondant = response_sheet.getRange(last_row,2).getValue();//最新の回答者のID（B列）を取得
  let counter_obj = {}; //カウント用の連想配列 {"選択肢":(応募数), ... }

  //シート「回答」を２～最終行目まで走査
  for(let i=2; i<=last_row; i++ ){
    let id = response_sheet.getRange(i,2).getValue(); //「社員IDを入力してください」回答
    let choice = response_sheet.getRange(i,3).getValue(); //「参加を希望する日程を選択してください」回答

    if(id.substr(0,6)=="[重複削除]"){ //文頭に[重複削除]=重複を計上対象外とする処理済のとき
      continue //計上しないため、次のループへ
    }
    if(i != last_row && id == latest_respondant){ //最新の回答以前に同一IDからの回答があり、かつ重複処理をまだ行っていない
      response_sheet.getRange(i,2).setValue("[重複削除]" + id); //文頭に[重複削除]
      response_sheet.getRange(i,3).setValue("[重複削除]" + choice); //文頭に[重複削除]
      continue //計上しないため、次のループへ
    }
    //送信された回答から、各選択肢をキーとして応募数を格納
    if( choice in counter_obj){ //連想配列に格納済の選択肢のとき
      counter_obj[choice] += 1; //カウントアップ
    }else{ //連想配列に未格納の選択肢のとき
      counter_obj[choice] = 1; //カウント開始
    }
  }
  return counter_obj;
}


//2.応募数の入力、受付可能な選択肢のリストを取得
function UpdateCapacityAndGetOptionList(counter_obj){
  let available_options = [];
  let last_row = capacity_sheet.getLastRow(); //最終行数
  //シート「各日程の定員管理」を２～最終行目まで走査
  for( let i=2; i<=last_row; i++){
    let choice = capacity_sheet.getRange(i,1).getValue(); //A列：選択肢
    let capacity = capacity_sheet.getRange(i,2).getValue(); //B列：定員
    
    if (choice in counter_obj){ //各選択肢に対応する応募数を、counter_objから取得
      var application_num = counter_obj[choice];
    }else{ //counter_objに格納なし=応募数ゼロ
      var application_num = 0;
    }
    capacity_sheet.getRange(i,3).setValue(application_num);//C列に応募数を入力
   
    //定員が応募を上回っているとき、受付可能な選択肢リストavailable_optionsにその選択肢を追加
    if(capacity - application_num > 0 ){available_options.push(choice)}
  }
  return available_options
}

//4.formの説明を作成
function CreateDescription(){
  const cnt_start = 5; //ここで設定した数値を残席が下回るとき、Formの説明に「選択肢　…　残りＸＸ名」と表示される
  description_text = "";
  let last_row = capacity_sheet.getLastRow(); //最終行数
  //シート「各日程の定員管理」を２～最終行目まで走査
  for(let i=2; i<=last_row; i++){
    //定員-応募数が残りX名以下の選択肢のみ対象に「選択肢　…　残りＸＸ名\n」という文言を作成
    let choice= capacity_sheet.getRange(i,1).getValue(); //A列：選択肢
    let capacity =  capacity_sheet.getRange(i,2).getValue(); //B列：定員
    let application_num = capacity_sheet.getRange(i,3).getValue(); //C列：応募数
    
    let left_num = capacity - application_num; //残り予約枠数
    if(left_num <= cnt_start){ //残り予約枠数が設定した数値（cnt_start）以下のとき
      let add_message = choice + " … 残り" + left_num + "名\n";
      description_text += add_message;
    }

  }
  return description_text
}


//----------------------------------------------------------------------------------
//任意のタイミングで実施
function SendRemindMail(){
  //フォームの回答シートから、回答のあった社員ＩＤをリストにして取得
  let respondar_list = [];
  let last_response = response_sheet.getLastRow(); //「フォームの回答1」最終行数
  for(let i =2; i<=last_row; i++){
    let staffID = response_sheet.getRange(i,2).getValue(); //社員IDの回答
    if (respondar_list.includes(staffIF)) reapondar_list.push(staffID); //社員IDをリストに格納（各１回）
  }
  //
  //現在の日付・時刻を取得
  const today = new Date();
  const year = today.getFullYear(); //年
  const month = today.getMonth()+1; //月*0～11
  const date = today.getDate(); //日
  const hr = today.getHours(); //時刻_時
  const mnt = ('0' + today.getMinutes()).slice(-2); //時刻_分 0埋めで二桁表示にする

  //メール文面に利用する現在日時+時刻
  const data_and_time = year + "年"+ month + "月" + date + "日" + hr + ":" + mnt;
  
  //メールの件名
  const subject = "【回答依頼】"+form_title+"に関するリマインド";

  const address_sheet = ss.getSheetByName("配信対象者")
  let last_row = address_sheet.getLastRow(); //「配信対象者」最終行数
  for(let i = 2; i<=last_row; i++){
    let targetID = address_sheet.getRange(i,1).getValue();
    if(respondar_list.includes(targetID)){
      //社員IDが回答済リストに存在する-->何もしない
    }else{ //社員IDが回答済みＩＤリストになければ
      let target_name = address_sheet.getRange(i,2).getValue();
      let target_emailaddress = address_sheet.getRange.getValue();

      let body =`${target_name} 様
      
      このメッセージは${data_and_time}時点で下記のフォームに未回答の社員へ送信されております。
      ${form_title}
      ${form_URL}
      
      申し込み可能日程には限りがあります。
      お早めにご検討願います。
      ご不明な点がございましたら下記担当までご連絡ください。
      
      総務部総務課　研修担当
      玉置（内線１０９８）  `

    }
    GmailApp.sendEmail(target_emailaddress, subject, body) //メール送信
  }


}


//----------------------------------------------------------------------------------
//初回のみ実施
//各日程の定員管理の「選択肢」を選択肢として反映する
function SetChoices(){
  //各日程の定員管理シートA列から選択肢を取得し、option_listに格納
  let option_list = [];
  let last_row = capacity_sheet.getLastRow();
  for (let i=2; i<=last_row; i++){
    let option = capacity_sheet.getRange(i,1).getValue();
    option_list.push(option);
  }
  //option_listを選択肢としてセット
  let questions = form.getItems(); //対象フォームの全質問取得
  let target_question = questions[1]; //2番目の質問を取得※0からのカウントで何番目か
  target_question.asMultipleChoiceItem().setChoiceValues(option_list); //選択肢をセット
}





