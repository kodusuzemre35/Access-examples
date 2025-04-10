<html lang="en" xmlns="http://www.w3.org/1999/xhtml">

<img align="left" src="Images/ReadMe/App.png" width="64px" >

# Microsoft Access Examples
Various examples of VBA, queries, macros, forms, reports and ribbon XML in an Microsoft Access database file

<!--[![Donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://www.paypal.me/AnthonyDuguid/1.00)-->
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE "MIT License Copyright © Anthony Duguid")
![current_build Office_2016](https://img.shields.io/badge/current_build-Office_2016-red.svg)
[![Latest Release](https://img.shields.io/github/release/Access-projects/Access-examples.svg?label=latest%20release)](https://github.com/Access-projects/Access-examples/releases)
[![Github commits (since latest release)](https://img.shields.io/github/commits-since/Access-projects/Access-examples/latest.svg)](https://github.com/Access-projects/Access-examples/commits/master)
[![GitHub issues](https://img.shields.io/github/issues/Access-projects/Access-examples.svg)](https://github.com/Access-projects/Access-examples/issues)
<!--[![Github All Releases](https://img.shields.io/github/downloads/Access-projects/Access-examples/total.svg)](https://github.com/Access-projects/Access-examples/releases)-->

## Table of Contents
 - <a href="#references">References</a>
 - <a href="#cmd-line">Command Line Options</a>
 - <a href="#object-list">Object Listing Reference</a>

<a id="user-content-references" class="anchor" href="#references" aria-hidden="true"> </a>
### References
|Link                        |Type                 |
|:-------------------------------|:--------------------------|
|[Microsoft Access Find & Replace Add-in](http://www.rickworld.com/products.html)|Software|
|[Microsoft Access Merge & Diff](http://www.accdbmerge.net/download)|Software|
|[O'Reilly Access Database Design & Programming, 3rd Edition](http://shop.oreilly.com/product/9780596002732.do)|Book|

<a id="user-content-cmd-line" class="anchor" href="#cmd-line" aria-hidden="true"> </a>

# 🤖 AI Project Management - Microsoft Access Sample Database

Bu proje, Microsoft Access ile geliştirilen bir **Yapay Zeka Proje Yönetim Veritabanı** örneğidir. VBA, SQL, makrolar, özel ribbon XML ve detaylı raporları içerir.

## 🗂️ İçerik Tablosu

- [Formlar](#formlar)
- [Makrolar](#makrolar)
- [Raporlar](#raporlar)
- [Modüller (VBA)](#modüller-vba)
- [Tablolar](#tablolar)
- [Sorgular (Queries)](#sorgular-queries)
- [Ribbon XML](#ribbon-xml)
- [Komut Satırı Seçenekleri](#komut-satırı-seçenekleri)

---

## 📋 Formlar

| ID   | Adı                             | Açıklama                                       |
|------|----------------------------------|------------------------------------------------|
| 001  | `Project_Overview_frm`          | Proje genel bilgileri                          |
| 002  | `Model_Selection_frm`           | Kullanılan AI modelleri                        |
| 003  | `Training_Config_frm`           | Eğitim parametrelerinin girildiği form         |
| 004  | `DataSet_Upload_frm`            | Veri seti yükleme formu                        |
| 005  | `User_Login_frm`                | Kullanıcı giriş ekranı                         |
| 006  | `Admin_Settings_frm`            | Yetkiler ve uygulama ayarları                  |
| 007  | `AI_Training_Status_frm`        | Eğitim durumu takibi                           |
| 008  | `Performance_Visualization_frm` | Başarı oranı ve metrik görselleştirmeleri     |
| 009  | `Error_Log_frm`                 | Eğitimde oluşan hataların listesi              |
| 010  | `Model_Versioning_frm`          | Model sürüm kontrol paneli                     |
| 011  | `Experiment_List_frm`           | AI deneyleri listesi                           |
| 012  | `Deployment_Log_frm`            | Yayınlanmış modellerin kaydı                   |
| 013  | `Documentation_frm`             | Proje açıklamaları ve dokümantasyon alanı      |

---

## ⚙️ Makrolar

| Makro Adı   | Açıklama                                        |
|-------------|-------------------------------------------------|
| `AutoExec`  | Uygulama başlatıldığında ilk çalışan makro      |
| `LogActivity` | Kullanıcı etkinliklerini kaydeder              |
| `Ribbon`    | Özel Access Ribbon XML menüsünü yükler          |

---

## 📊 Raporlar

| ID   | Adı                          | Açıklama                                      |
|------|-----------------------------|-----------------------------------------------|
| 001  | `AI_Project_Summary_rpt`    | Proje genel özet raporu                       |
| 002  | `Model_Performance_rpt`     | Modellerin başarı oranları                    |
| 003  | `Training_Log_rpt`          | Eğitim süreçlerinin detaylı log'u             |
| 004  | `Deployment_Timeline_rpt`   | Zaman çizelgesi ile yayınlama geçmişi         |
| 005  | `Error_Report_rpt`          | Hata kayıtları ve detayları                   |
| 006  | `Data_Distribution_rpt`     | Veri seti analizleri                          |
| 007  | `Feature_Usage_rpt`         | Kullanılan özelliklerin sıklığı               |
| 008  | `User_Activity_Log_rpt`     | Kullanıcı etkinlik geçmişi                    |

---

## 💻 Modüller (VBA)

| Modül Adı          | Açıklama                                                 |
|--------------------|----------------------------------------------------------|
| `modAI_Math`        | AI ile ilgili temel hesaplamalar ve istatistik fonksiyonları |
| `modTrainControl`   | Model eğitim süreci kontrol fonksiyonları               |
| `modFileManager`    | Dosya yönetimi, veri yükleme ve tarama işlemleri        |
| `modGraph`          | Eğitim süreci grafiklerini oluşturma                    |
| `modUserControl`    | Giriş/çıkış, kullanıcı kontrol işlemleri                |
| `modDataPrep`       | Veri temizleme, ön işleme                               |
| `modErrorHandler`   | Hata loglama ve yönetim fonksiyonları                   |
| `modAPI_Integration`| API bağlantıları (OpenAI, HuggingFace vb.)             |
| `modMetrics`        | Precision, Recall, F1 Score, AUC gibi AI metrikleri     |

---

## 🗃️ Tablolar

- `tblProjects`
- `tblModels`
- `tblTrainings`
- `tblDatasets`
- `tblUsers`
- `tblLogs`
- `tblErrors`
- `tblDeployments`
- `tblHyperparameters`
- `tblFeatureImportance`

---

## 🔍 Sorgular (Queries)

- `qry_ModelSuccessRates`
- `qry_TrainingByDatasetSize`
- `qry_UserLoginActivity`
- `qry_HyperparameterUsage`
- `qry_MonthlyModelDeployments`
- `qry_ErrorStatistics`
- `qry_FeatureUsageStats`

---

## 🧾 Ribbon XML

Bu veritabanında özel bir Ribbon XML kullanılır. "Proje Yönetimi", "Model Eğitimi", "Yayınlama", "Raporlar" gibi özel sekmeler içerir.

---

## 🖥 Komut Satırı Seçenekleri (Access)

| Seçenek | Açıklama |
|---------|----------|
| `/decompile` | Eski kodları temizler, performansı artırabilir |
| `/excl` | Veritabanını yalnızca sizin için açar |
| `/ro` | Salt okunur olarak açar |
| `/pwd` | Şifre ile veritabanı açar |
| `/compact` | Veritabanını sıkıştırır ve onarır |
| `/x` | Belirtilen makroyu çalıştırır |
| `/cmd
