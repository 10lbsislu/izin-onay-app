# 📋 Teams İzin Onay Sistemi

Microsoft Teams ile tam entegre, kurumsal izin talep ve onay uygulaması.

## ✨ Özellikler

| Çalışan | Yönetici |
|---------|----------|
| İzin talebi oluşturma | Bekleyen talepleri görme |
| Başlangıç/bitiş tarihi seçimi | Tek tıkla onayla / reddet |
| İzin türü seçimi (yıllık, hastalık, mazeret...) | Yorum/gerekçe ekleme |
| Talep geçmişini görme | Tüm talepler tablosu + filtreleme |
| Durum takibi (Beklemede / Onaylandı / Reddedildi) | **Kuruluştan onaylayıcı ekleme** (adı yazınca çıkar) |
| Teams'de otomatik bildirim alma | Kanala otomatik Adaptive Card |

## 🏗️ Teknolojiler

- **Frontend**: React 18 + TypeScript + Fluent UI v9
- **Backend**: Azure Functions v4 (Node.js 20)
- **Veri**: SharePoint Excel (Graph Workbook API)
- **Auth**: Azure AD SSO + OBO Flow
- **Bildirim**: Teams Adaptive Cards + Direkt Mesaj

## 🚀 Hızlı Başlangıç

```bash
# Bağımlılıkları yükle
npm install

# .env dosyasını oluştur
cp .env.example .env
# Değerleri doldur

# Geliştirme sunucusunu başlat
npm run dev
```

Tam kurulum için [setup/SETUP.md](setup/SETUP.md) dosyasını inceleyin.

## 📁 Proje Yapısı

```
izin-onay-app/
├── src/
│   ├── components/
│   │   ├── RequestForm.tsx      # İzin talebi formu
│   │   ├── MyRequests.tsx       # Çalışan talep listesi
│   │   ├── AdminPanel.tsx       # Yönetici paneli
│   │   ├── UserPicker.tsx       # Kuruluş kullanıcı arama
│   │   └── StatusBadge.tsx      # Durum rozeti
│   ├── services/
│   │   └── requestService.ts    # API çağrıları
│   ├── hooks/
│   │   └── useTeamsContext.ts   # Teams SSO hook
│   ├── auth/
│   │   └── authConfig.ts        # MSAL yapılandırması
│   └── types/index.ts           # TypeScript tipleri
├── api/
│   ├── getRequests/             # GET talepleri
│   ├── createRequest/           # POST yeni talep
│   ├── updateRequest/           # PUT onayla/reddet
│   ├── searchUsers/             # GET org kullanıcı arama
│   ├── getApprovers/            # GET onaylayıcılar
│   ├── manageApprovers/         # POST ekle/kaldır
│   └── shared/
│       ├── graphClient.js       # Graph API istemcisi
│       ├── excelService.js      # Excel okuma/yazma
│       └── notifyService.js     # Adaptive Card bildirimleri
├── manifest/
│   └── manifest.json            # Teams app manifest
└── setup/
    └── SETUP.md                 # Kurulum rehberi
```
