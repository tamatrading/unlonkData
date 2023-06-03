# これはサンプルの Python スクリプトです。

# Shift+F10 を押して実行するか、ご自身のコードに置き換えてください。
# Shift を2回押す を押すと、クラス/ファイル/ツールウィンドウ/アクション/設定を検索します。


from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

def get_project_names():
    # Setup driver
    s = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=s)

    # Open the webpage
    driver.get('https://token.unlocks.app/')

    # Find project names
    project_elements = driver.find_elements(By.XPATH, '//*[@id="__next"]/div[1]/div[2]/div[2]/div[3]/div[2]/table/tbody/tr')

    cnt = 0
    # Extract and print project names
    for prj in project_elements:
        unlockName = prj.find_element(By.XPATH, 'th/div/a/div/div[2]/p').text     #PROJECT NAME
        unlockPer = prj.find_element(By.XPATH, 'td[7]/a/div/div/div[1]/div[2]/p[1]').text       #アンロック割合(もしくは量)
        unlockDate = prj.find_element(By.XPATH, 'td[7]/a/div/div/div[1]/div[6]/p[1]').text      #アンロック日付
        unlockTime = prj.find_element(By.XPATH, 'td[7]/a/div/div/div[1]/div[6]/p[2]').text      #アンロック時刻(JST)

        print(f"{unlockName} : {unlockPer} : {unlockDate} : {unlockTime}")
        cnt = cnt + 1
        if cnt >= 20:
            break

    # Close the driver
    driver.quit()

if __name__ == '__main__':
    get_project_names()
