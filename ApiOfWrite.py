#查询系列api，每6小时调用一次（周六日不启动）
name: Run Api.Write

on: 
  release:
    types: [published]
  push:
    tags:
    - 'v*'
  #  branches: 
  #    - master
  schedule:
    - cron: '12 */12 * * 1-5'
  watch:
    types: [started]
  workflow_dispatch:
   
jobs:
  build:
    runs-on: ubuntu-latest
    #if: github.event.repository.owner.id == github.event.sender.id  # 自己点的 start
    steps:
    - name: Checkout
      uses: actions/checkout@master
    - name: Install requests #安装requests模块
      run: |
        pip install requests XlsxWriter
    - name: run api.Write #api调用
      env: 
        CITY: ${{ secrets.CITY }}
        #github的账号信息
        GH_TOKEN: ${{ secrets.GH_TOKEN }} 
        GH_REPO: ${{ github.repository }}
        #以下是微软的账号信息(修改以下，复制增加账号)
        APP_NUM: ${{ secrets.APP_NUM }} 
        #账号/应用1
        MS_TOKEN_1: ${{ secrets.MS_TOKEN }} 
        CLIENT_ID_1: ${{ secrets.CLIENT_ID }}
        CLIENT_SECRET_1: ${{ secrets.CLIENT_SECRET }}
        #账号/应用2
        MS_TOKEN_2: ${{ secrets.MS_TOKEN_2 }} 
        CLIENT_ID_2: ${{ secrets.CLIENT_ID_2 }}
        CLIENT_SECRET_2: ${{ secrets.CLIENT_SECRET_2 }}
        #账号/应用3
        MS_TOKEN_3: ${{ secrets.MS_TOKEN_3 }} 
        CLIENT_ID_3: ${{ secrets.CLIENT_ID_3 }}
        CLIENT_SECRET_3: ${{ secrets.CLIENT_SECRET_3 }}
        #账号/应用4
        MS_TOKEN_4: ${{ secrets.MS_TOKEN_4 }} 
        CLIENT_ID_4: ${{ secrets.CLIENT_ID_4 }}
        CLIENT_SECRET_4: ${{ secrets.CLIENT_SECRET_4 }}
        #账号/应用5
        MS_TOKEN_5: ${{ secrets.MS_TOKEN_5 }} 
        CLIENT_ID_5: ${{ secrets.CLIENT_ID_5 }}
        CLIENT_SECRET_5: ${{ secrets.CLIENT_SECRET_5 }}
        #如此类推，自己复制增加账号/应用
        MS_TOKEN_n: ${{ secrets.MS_TOKEN_n }} 
        CLIENT_ID_n: ${{ secrets.CLIENT_ID_n }}
        CLIENT_SECRET_n: ${{ secrets.CLIENT_SECRET_n }}
      run: |
        python ApiOfWrite.py
