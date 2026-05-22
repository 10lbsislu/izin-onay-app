# 🚀 Kurulum Rehberi — Teams İzin Onay Sistemi

Bu rehber uygulamayı sıfırdan prod'a taşımanın tüm adımlarını içerir.

---

## 1. Azure AD Uygulama Kaydı

### 1.1 App Registration oluşturun
1. [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Ad: `IzinOnayApp`
3. Desteklenen hesap türleri: **Bu kuruluşun hesapları**
4. Redirect URI: `https://localhost:3000` (geliştirme için)
5. **Register** butonuna tıklayın

### 1.2 API izinleri ekleyin
**API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**:

| İzin | Açıklama |
|------|----------|
| `User.Read` | Giriş yapan kullanıcıyı okuma |
| `User.ReadBasic.All` | Kuruluş kullanıcılarını arama |
| `Files.ReadWrite` | SharePoint Excel okuma/yazma |
| `ChannelMessage.Send` | Kanala mesaj gönderme |
| `Chat.ReadWrite` | Kişiye direkt mesaj |

**Microsoft Graph** → **Application** (servis hesabı için):

| İzin | Açıklama |
|------|----------|
| `User.ReadBasic.All` | Kullanıcı arama |
| `Files.ReadWrite.All` | Excel erişimi |
| `ChannelMessage.Send` | Kanal mesajı |
| `Chat.Create` | Yeni chat oluşturma |

→ **Grant admin consent** butonuna tıklayın ✅

### 1.3 Client Secret oluşturun
**Certificates & secrets** → **New client secret**
- Açıklama: `IzinOnaySecret`
- Son kullanma: 24 ay
- **Value** değerini kopyalayın (bir daha görünmez!)

### 1.4 Expose an API (Teams SSO için)
**Expose an API** → **Set** → Application ID URI:
```
api://YOUR_DOMAIN/YOUR_CLIENT_ID
```

**Add a scope:**
- Scope name: `access_as_user`
- Who can consent: Admins and users
- Display name: `Access as user`

**Add a client application:**
- Teams Desktop: `1fec8e78-bce4-4aaf-ab1b-5451cc387264`
- Teams Web: `5e3ce6c0-2b1f-4285-8d4b-75ee78787346`

---

## 2. SharePoint Excel Dosyasını Hazırlayın

### 2.1 Excel dosyası oluşturun
1. SharePoint sitenize gidin
2. **Belgeler** klasöründe **IzinTalepleri.xlsx** oluşturun

### 2.2 "Talepler" sayfası
A1'den itibaren şu başlıkları girin:
```
id | requesterId | requesterName | requesterEmail | leaveType | startDate | endDate | totalDays | description | status | approverId | approverName | approverComment | createdAt | updatedAt
```
Tüm sütunları seçin → **Insert** → **Table** → "My table has headers" ✓ → Tablo adı: `TaleplerTablosu`

### 2.3 "Onaylayıcılar" sayfası
A1'den:
```
id | displayName | mail | addedAt
```
Tablo adı: `OnaylayıcılarTablosu`

### 2.4 SharePoint Site ID ve Dosya ID'sini bulun
Browser'da şu URL'yi açın:
```
https://graph.microsoft.com/v1.0/sites/{TENANT}.sharepoint.com:/sites/{SITE_ADI}
```
→ `id` alanını kopyalayın → `SHAREPOINT_SITE_ID`

```
https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive/root:/IzinTalepleri.xlsx
```
→ `id` alanını kopyalayın → `EXCEL_FILE_ID`

---

## 3. Teams Kanal ID Bulma

1. Teams'de hedef kanala sağ tıklayın → **Get link to channel**
2. URL'den `groupId` ve `channel ID` değerlerini alın:
   ```
   https://teams.microsoft.com/.../channel/19:xxx@thread.tacv2/...
   ```
   - `groupId` → `TEAMS_TEAM_ID`
   - `19:xxx@thread.tacv2` → `TEAMS_CHANNEL_ID`

---

## 4. Azure Functions Deploy

### 4.1 Azure Functions App oluşturun
```bash
az login
az group create --name IzinOnayRG --location westeurope
az storage account create --name izinstore --resource-group IzinOnayRG --sku Standard_LRS
az functionapp create \
  --name izin-onay-api \
  --resource-group IzinOnayRG \
  --storage-account izinstore \
  --consumption-plan-location westeurope \
  --runtime node \
  --runtime-version 20
```

### 4.2 Environment değişkenlerini ayarlayın
```bash
az functionapp config appsettings set \
  --name izin-onay-api \
  --resource-group IzinOnayRG \
  --settings \
    TENANT_ID="xxx" \
    CLIENT_ID="xxx" \
    CLIENT_SECRET="xxx" \
    SHAREPOINT_SITE_ID="xxx" \
    EXCEL_FILE_ID="xxx" \
    TEAMS_TEAM_ID="xxx" \
    TEAMS_CHANNEL_ID="xxx" \
    FRONTEND_URL="https://your-frontend.azurestaticapps.net"
```

### 4.3 Deploy edin
```bash
cd api
npm install
func azure functionapp publish izin-onay-api
```

---

## 5. Frontend Deploy (Azure Static Web Apps)

```bash
az staticwebapp create \
  --name izin-onay-frontend \
  --resource-group IzinOnayRG \
  --location westeurope \
  --source . \
  --branch main \
  --app-location "/" \
  --output-location "dist"
```

`.env` dosyasını prod için doldurun:
```env
VITE_CLIENT_ID=your-client-id
VITE_TENANT_ID=your-tenant-id
VITE_REDIRECT_URI=https://izin-onay-frontend.azurestaticapps.net
VITE_API_BASE_URL=https://izin-onay-api.azurewebsites.net/api
```

Build ve deploy:
```bash
npm install
npm run build
```

---

## 6. Teams Manifest'i Yükleyin

### 6.1 Manifest'i güncelleyin
`manifest/manifest.json` içindeki placeholder'ları doldurun:
- `{{TEAMS_APP_ID}}` → yeni bir GUID (örn. `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`)
- `{{FRONTEND_DOMAIN}}` → Azure Static Web App domain'i
- `{{CLIENT_ID}}` → Azure AD Client ID

### 6.2 İkon dosyaları ekleyin
- `manifest/color.png` — 192×192 px renkli ikon
- `manifest/outline.png` — 32×32 px beyaz outline ikon

### 6.3 Zip ve yükle
```bash
cd manifest
zip -r ../izin-onay-app.zip manifest.json color.png outline.png
```

Teams Admin Center → **Manage apps** → **Upload new app** → zip dosyasını seçin

---

## 7. Uygulama Kurulumu (İlk Çalıştırma)

1. Teams'de uygulamayı bir kanala ekleyin
2. Yönetici hesabıyla giriş yapın
3. **Yönetici Paneli** → **Onaylayıcı Yönetimi** sekmesine gidin
4. Onaylayıcı olarak kendinizi veya yöneticileri ekleyin
5. **Ayarlar** bölümünden bildirim kanalını seçin

---

## 8. Geliştirme Ortamı

```bash
# Frontend
cp .env.example .env
# .env'i doldurun
npm install
npm run dev

# API (ayrı terminal)
cd api
cp local.settings.json.example local.settings.json
# local.settings.json'u doldurun
npm install
func start
```

---

## Mimari Özeti

```
Teams Tab (React + Fluent UI)
        │
        ├── Teams SSO → Bearer Token
        │
        ▼
Azure Functions API
        │
        ├── Graph API (OBO Flow)
        │       ├── /users?$search → Kullanıcı arama
        │       ├── SharePoint Excel → Veri saklama
        │       └── Teams Channel/Chat → Bildirimler
        │
        └── Excel (SharePoint)
                ├── TaleplerTablosu
                └── OnaylayıcılarTablosu
```

---

## Destek

Sorun yaşarsanız şu logları kontrol edin:
- **Azure Functions**: Azure Portal → Function App → Monitor
- **Frontend**: Browser DevTools → Console
- **Graph API**: [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)
