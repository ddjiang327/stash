import React, { useState, useEffect } from 'react';
import {
  View,
  Text,
  TextInput,
  TouchableOpacity,
  ScrollView,
  Image,
  StyleSheet,
  Modal,
  Alert,
  ActivityIndicator,
  KeyboardAvoidingView,
  Platform,
  StatusBar,
  FlatList,
  Pressable,
} from 'react-native';
import AsyncStorage from '@react-native-async-storage/async-storage';
import * as ImagePicker from 'expo-image-picker';
import { manipulateAsync, SaveFormat } from 'expo-image-manipulator';
import * as FileSystem from 'expo-file-system/legacy';
import * as Sharing from 'expo-sharing';
import { Ionicons } from '@expo/vector-icons';
import { SafeAreaProvider, useSafeAreaInsets } from 'react-native-safe-area-context';
import * as XLSX from 'xlsx';
import * as Localization from 'expo-localization';

// ─── i18n ─────────────────────────────────────────────────────────────────────
const TRANSLATIONS = {
  zh: {
    appName: '物归处',
    itemCount: (n) => `共 ${n} 件`,
    foundCount: (n) => `找到 ${n} 件`,
    searchPlaceholder: '搜索物品名称或位置…',
    all: '全部',
    manage: '+ 管理',
    emptySearch: '没有找到匹配的物品',
    emptyList: '还没有记录，点击 + 开始添加',
    addItem: '添加物品',
    editItem: '编辑物品',
    save: '保存',
    itemPhoto: '物品照片',
    itemPhotoHint: '拍摄或上传物品照片',
    itemName: '物品名称 *',
    itemNamePlaceholder: '例：护照、电钻、急救箱…',
    location: '存放位置 *',
    locationPlaceholder: '例：书房第二层书架左侧',
    locationPhoto: '位置照片（可选）',
    locationPhotoHint: '拍摄存放位置的照片',
    category: '分类',
    savedToast: (name) => `${name} 已保存 ✓`,
    deletedToast: '已删除',
    backupToast: '备份已分享 ✓',
    xlsxToast: '表格已分享 ✓',
    exportXlsx: '导出表格',
    exportJson: '备份数据',
    exportTitle: '导出',
    colName: '物品名称', colLocation: '存放位置', colCategory: '分类', colDate: '记录日期',
    storedAt: '存放位置',
    locationPhotoLabel: '位置照片',
    recordedOn: (d) => `记录于 ${d}`,
    confirmDelete: '确认删除这件物品？',
    cancel: '取消',
    delete: '删除',
    manageCategories: '管理分类',
    newCategoryPlaceholder: '新分类名称',
    add: '添加',
    notRecorded: '未记录',
    noLocation: '未填位置',
    compressing: '压缩中…',
    permissionTitle: '需要权限',
    permissionMsg: '需要访问相册权限才能添加照片',
    defaultCats: ['文件证件','电子设备','工具五金','衣物配饰','药品保健','厨房用品','书籍文具','帮人保管','其他'],
    dateLocale: 'zh-CN',
    takePhoto: '📷 拍照',
    chooseFromLibrary: '🖼️ 从相册选择',
    cameraPermissionMsg: '需要相机权限才能拍照',
  },
  en: {
    appName: 'Cached',
    itemCount: (n) => `${n} item${n !== 1 ? 's' : ''}`,
    foundCount: (n) => `${n} found`,
    searchPlaceholder: 'Search by name or location…',
    all: 'All',
    manage: '+ Manage',
    emptySearch: 'No items match your search',
    emptyList: 'Nothing here yet — tap + to add',
    addItem: 'Add Item',
    editItem: 'Edit Item',
    save: 'Save',
    itemPhoto: 'Item Photo',
    itemPhotoHint: 'Take or upload a photo',
    itemName: 'Item Name *',
    itemNamePlaceholder: 'e.g. Passport, Drill, First Aid Kit…',
    location: 'Storage Location *',
    locationPlaceholder: 'e.g. Office shelf, second row left',
    locationPhoto: 'Location Photo (optional)',
    locationPhotoHint: "Photo of where it's stored",
    category: 'Category',
    savedToast: (name) => `${name} saved ✓`,
    deletedToast: 'Deleted',
    backupToast: 'Backup shared ✓',
    xlsxToast: 'Spreadsheet shared ✓',
    exportXlsx: 'Export to Excel',
    exportJson: 'Backup data',
    exportTitle: 'Export',
    colName: 'Item Name', colLocation: 'Location', colCategory: 'Category', colDate: 'Date recorded',
    storedAt: 'Stored At',
    locationPhotoLabel: 'Location Photo',
    recordedOn: (d) => `Recorded on ${d}`,
    confirmDelete: 'Delete this item?',
    cancel: 'Cancel',
    delete: 'Delete',
    manageCategories: 'Manage Categories',
    newCategoryPlaceholder: 'New category name',
    add: 'Add',
    notRecorded: 'Not recorded',
    noLocation: 'No location set',
    compressing: 'Compressing…',
    permissionTitle: 'Permission needed',
    permissionMsg: 'We need photo library access to add photos',
    defaultCats: ['Documents','Electronics','Tools','Clothing','Medicine','Kitchen','Books','Keeping for Others','Other'],
    dateLocale: 'en-AU',
    takePhoto: '📷 Take Photo',
    chooseFromLibrary: '🖼️ Choose from Library',
    cameraPermissionMsg: 'Camera permission is needed to take photos',
  },
  fr: {
    appName: 'Cached',
    itemCount: (n) => `${n} objet${n !== 1 ? 's' : ''}`,
    foundCount: (n) => `${n} trouvé${n !== 1 ? 's' : ''}`,
    searchPlaceholder: 'Rechercher par nom ou emplacement…',
    all: 'Tout',
    manage: '+ Gérer',
    emptySearch: 'Aucun résultat',
    emptyList: 'Rien ici — appuyez sur + pour ajouter',
    addItem: 'Ajouter',
    editItem: 'Modifier',
    save: 'Enregistrer',
    itemPhoto: "Photo de l'objet",
    itemPhotoHint: 'Prendre ou importer une photo',
    itemName: "Nom de l'objet *",
    itemNamePlaceholder: 'ex : Passeport, Perceuse, Trousse…',
    location: 'Emplacement *',
    locationPlaceholder: 'ex : Étagère bureau, 2e rangée gauche',
    locationPhoto: "Photo de l'emplacement (optionnel)",
    locationPhotoHint: "Photo de l'endroit où c'est rangé",
    category: 'Catégorie',
    savedToast: (name) => `${name} enregistré ✓`,
    deletedToast: 'Supprimé',
    backupToast: 'Sauvegarde partagée ✓',
    xlsxToast: 'Tableau partagé ✓',
    exportXlsx: 'Exporter en Excel',
    exportJson: 'Sauvegarder les données',
    exportTitle: 'Exporter',
    colName: 'Nom', colLocation: 'Emplacement', colCategory: 'Catégorie', colDate: 'Date enregistrée',
    storedAt: 'Rangé à',
    locationPhotoLabel: "Photo de l'emplacement",
    recordedOn: (d) => `Enregistré le ${d}`,
    confirmDelete: 'Supprimer cet objet ?',
    cancel: 'Annuler',
    delete: 'Supprimer',
    manageCategories: 'Gérer les catégories',
    newCategoryPlaceholder: 'Nom de la nouvelle catégorie',
    add: 'Ajouter',
    notRecorded: 'Non renseigné',
    noLocation: 'Emplacement non renseigné',
    compressing: 'Compression…',
    permissionTitle: 'Permission requise',
    permissionMsg: "Nous avons besoin d'accéder à votre galerie",
    defaultCats: ['Documents','Électronique','Outils','Vêtements','Médicaments','Cuisine','Livres',"Garde d'objets",'Autre'],
    dateLocale: 'fr-FR',
    takePhoto: '📷 Prendre une photo',
    chooseFromLibrary: '🖼️ Choisir dans la galerie',
    cameraPermissionMsg: "L'autorisation d'accès à la caméra est nécessaire pour prendre des photos",
  },
};

// ─── Presets ──────────────────────────────────────────────────────────────────
const PRESETS = {
  zh: {
    docs:        ['护照','身份证','驾照','户口本','结婚证','出生证明','毕业证','房产证','保险单','合同'],
    electronics: ['手机','平板','笔记本电脑','充电器','耳机','相机','移动硬盘','U盘','路由器','延长线'],
    tools:       ['电钻','螺丝刀','锤子','扳手','卷尺','水平仪','美工刀','胶枪','砂纸','电工胶带'],
    clothing:    ['西装','冬季外套','礼服','泳衣','运动服','皮带','领带','围巾','手套','帽子'],
    medicine:    ['急救箱','血压计','体温计','血糖仪','常备药','创可贴','消毒水','口罩','维生素','眼药水'],
    kitchen:     ['锅','刀具','砧板','烤箱','料理机','保鲜盒','保温壶','咖啡机','筷子套装','计时器'],
    books:       ['字典','相册','日记本','教材','说明书','笔记本','文件夹','印章','邮票','地图'],
    other:       ['钥匙','备用钥匙','礼品卡','现金','充电宝','雨伞','手电筒','电池','胶水','剪刀'],
    keeping:     ['朋友的书','朋友的充电器','朋友的钥匙','借来的工具','暂存物品'],
  },
  en: {
    docs:        ['Passport',"Driver's License",'ID Card','Birth Certificate','Marriage Certificate','Insurance Policy','Lease Agreement','Tax Documents','Warranty Cards','Medical Records'],
    electronics: ['Phone','Tablet','Laptop','Charger','Earphones','Camera','External Hard Drive','USB Drive','Router','Extension Cord'],
    tools:       ['Drill','Screwdriver Set','Hammer','Wrench','Tape Measure','Level','Utility Knife','Glue Gun','Sandpaper','Electrical Tape'],
    clothing:    ['Suit','Winter Coat','Formal Dress','Swimwear','Sportswear','Belt','Tie','Scarf','Gloves','Hat'],
    medicine:    ['First Aid Kit','Blood Pressure Monitor','Thermometer','Glucose Meter','Daily Meds','Bandages','Antiseptic','Face Masks','Vitamins','Eye Drops'],
    kitchen:     ['Pots & Pans','Knife Set','Cutting Board','Oven','Blender','Food Containers','Thermos','Coffee Maker','Chopsticks','Timer'],
    books:       ['Dictionary','Photo Album','Journal','Textbooks','Manuals','Notebook','Binder','Stamps','Maps','Receipts'],
    other:       ['Keys','Spare Keys','Gift Cards','Cash','Power Bank','Umbrella','Flashlight','Batteries','Glue','Scissors'],
    keeping:     ["Friend's Book","Friend's Charger","Friend's Keys",'Borrowed Tools','Items in Custody'],
  },
  fr: {
    docs:        ['Passeport','Permis de conduire',"Carte d'identité",'Acte de naissance','Acte de mariage','Assurance','Bail','Documents fiscaux','Garanties','Dossier médical'],
    electronics: ['Téléphone','Tablette','Ordinateur portable','Chargeur','Écouteurs','Appareil photo','Disque dur externe','Clé USB','Routeur','Multiprise'],
    tools:       ['Perceuse','Tournevis','Marteau','Clé à molette','Mètre ruban','Niveau','Cutter','Pistolet à colle','Papier de verre','Chatterton'],
    clothing:    ['Costume',"Manteau d'hiver","Robe de soirée","Maillot de bain","Tenue de sport",'Ceinture','Cravate','Écharpe','Gants','Chapeau'],
    medicine:    ['Trousse de secours','Tensiomètre','Thermomètre','Glucomètre','Médicaments','Pansements','Antiseptique','Masques','Vitamines','Gouttes oculaires'],
    kitchen:     ['Casseroles','Couteaux','Planche à découper','Four','Mixeur','Boîtes hermétiques','Thermos','Cafetière','Baguettes','Minuteur'],
    books:       ['Dictionnaire','Album photo','Journal intime','Manuels','Notices','Carnet','Classeur','Timbres','Cartes','Reçus'],
    other:       ['Clés','Clés de rechange','Cartes cadeaux','Espèces','Batterie externe','Parapluie','Lampe torche','Piles','Colle','Ciseaux'],
    keeping:     ["Livre d'un ami","Chargeur d'un ami","Clés d'un ami",'Outils empruntés','Objets en garde'],
  },
};

// ─── Constants ────────────────────────────────────────────────────────────────
const STORAGE_KEY     = 'cached_items_v1';
const CUSTOM_CATS_KEY = 'cached_custom_cats_v1';
const LANG_KEY        = 'cached_lang';
const DEFAULT_CAT_IDS = ['docs','electronics','tools','clothing','medicine','kitchen','books','keeping','other'];
const CAT_EMOJIS      = ['📄','💻','🔧','👗','💊','🍳','📚','🤝','📦'];

// ─── Colors ───────────────────────────────────────────────────────────────────
const C = {
  bg:        '#F7F5F0',
  surface:   '#FFFFFF',
  surface2:  '#EEEBE4',
  border:    '#E2DDD6',
  text:      '#1A1714',
  textMuted: '#8C8479',
  accent:    '#C84B2F',
  accentBg:  '#FDF0EC',
  danger:    '#E53935',
};

// ─── Helpers ──────────────────────────────────────────────────────────────────
function genId() { return Date.now().toString(36) + Math.random().toString(36).slice(2); }

function getDefaultCats(lang) {
  return DEFAULT_CAT_IDS.map((id, i) => ({
    id, emoji: CAT_EMOJIS[i], label: TRANSLATIONS[lang].defaultCats[i],
  }));
}

function getAllCats(lang, customCats) {
  return [...getDefaultCats(lang), ...customCats];
}

// ─── PhotoUpload Component ────────────────────────────────────────────────────
function PhotoUpload({ value, onChange, label, hint, compressingLabel, t }) {
  const [compressing, setCompressing] = useState(false);

async function handlePick() {
  Alert.alert(
    label,
    '',
    [
      {
        text: t.takePhoto,
        onPress: async () => {
          const { status } = await ImagePicker.requestCameraPermissionsAsync();
          if (status !== 'granted') {
            Alert.alert(t.permissionTitle, t.cameraPermissionMsg);
            return;
          }
          const result = await ImagePicker.launchCameraAsync({
            allowsEditing: true,
            quality: 1,
          });
          await processResult(result);
        },
      },
      {
        text: t.chooseFromLibrary,
        onPress: async () => {
          const { status } = await ImagePicker.requestMediaLibraryPermissionsAsync();
          if (status !== 'granted') {
            Alert.alert(t.permissionTitle, t.permissionMsg);
            return;
          }
          const result = await ImagePicker.launchImageLibraryAsync({
            mediaTypes: ['images'],
            allowsEditing: true,
            quality: 1,
          });
          await processResult(result);
        },
      },
      { text: t.cancel, style: 'cancel' },
    ]
  );
}

async function processResult(result: ImagePicker.ImagePickerResult) {
  if (!result.canceled && result.assets?.[0]) {
    setCompressing(true);
    const asset = result.assets[0];
    const MAX = 400;
    let w = asset.width || 800;
    let h = asset.height || 800;
    if (w > h) { if (w > MAX) { h = Math.round(h * MAX / w); w = MAX; } }
    else       { if (h > MAX) { w = Math.round(w * MAX / h); h = MAX; } }
    try {
      const res = await manipulateAsync(
        asset.uri,
        [{ resize: { width: w, height: h } }],
        { compress: 0.6, format: SaveFormat.JPEG, base64: true }
      );
      onChange(`data:image/jpeg;base64,${res.base64}`);
    } catch (e) {
      Alert.alert('Error', String(e));
    }
    setCompressing(false);
  }
}

  return (
    <View style={s.photoWrap}>
      <Text style={s.photoLabel}>{label}</Text>
      {compressing ? (
        <View style={s.photoPlaceholder}>
          <ActivityIndicator color={C.accent} />
          <Text style={[s.photoHint, { marginTop: 6 }]}>{compressingLabel}</Text>
        </View>
      ) : value ? (
        <View style={s.photoPreviewWrap}>
          <Image source={{ uri: value }} style={s.photoPreview} resizeMode="cover" />
          <TouchableOpacity onPress={() => onChange(null)} style={s.photoRemoveBtn}>
            <Ionicons name="close" size={14} color="#fff" />
          </TouchableOpacity>
        </View>
      ) : (
        <TouchableOpacity onPress={handlePick} style={s.photoPlaceholder}>
          <Ionicons name="camera-outline" size={22} color={C.textMuted} />
          <Text style={[s.photoHint, { marginTop: 6 }]}>{hint}</Text>
        </TouchableOpacity>
      )}
    </View>
  );
}

// ─── LangSwitcher Component ───────────────────────────────────────────────────
function LangSwitcher({ lang, setLang }) {
  const [open, setOpen] = useState(false);
  const langs = [
    { code: 'zh', flag: '🇨🇳', label: '中文' },
    { code: 'en', flag: '🇬🇧', label: 'English' },
    { code: 'fr', flag: '🇫🇷', label: 'Français' },
  ];
  const current = langs.find(l => l.code === lang);
  return (
    <View>
      <TouchableOpacity onPress={() => setOpen(true)} style={[s.iconBtn, { paddingHorizontal: 6, width: 'auto', minWidth: 40 }]}>
        <Ionicons name="globe-outline" size={16} color={C.textMuted} />
        <Text style={{ fontSize: 15, marginLeft: 3 }}>{current.flag}</Text>
      </TouchableOpacity>
      <Modal visible={open} transparent animationType="fade" onRequestClose={() => setOpen(false)}>
        <Pressable style={{ flex: 1 }} onPress={() => setOpen(false)}>
          <View style={s.langMenuOverlay}>
            <View style={s.langMenu}>
              {langs.map(l => (
                <TouchableOpacity key={l.code} onPress={() => { setLang(l.code); setOpen(false); }}
                  style={[s.langItem, l.code === lang && s.langItemActive]}>
                  <Text style={{ fontSize: 20 }}>{l.flag}</Text>
                  <Text style={[s.langItemText, l.code === lang && { color: C.accent, fontWeight: '700' }]}>{l.label}</Text>
                </TouchableOpacity>
              ))}
            </View>
          </View>
        </Pressable>
      </Modal>
    </View>
  );
}

// ─── BottomModal Component ────────────────────────────────────────────────────
function BottomModal({ visible, onClose, title, children }) {
  return (
    <Modal visible={visible} transparent animationType="slide" onRequestClose={onClose}>
      <View style={{ flex: 1 }}>
        <Pressable style={{ flex: 1, backgroundColor: 'rgba(0,0,0,0.45)' }} onPress={onClose} />
        <View style={s.bottomSheet}>
          <View style={s.bottomSheetHandle} />
          <View style={s.bottomSheetHeader}>
            <Text style={s.bottomSheetTitle}>{title}</Text>
            <TouchableOpacity onPress={onClose}>
              <Ionicons name="close" size={20} color={C.textMuted} />
            </TouchableOpacity>
          </View>
          <ScrollView showsVerticalScrollIndicator={false}>
            {children}
          </ScrollView>
        </View>
      </View>
    </Modal>
  );
}

// ─── ConfirmModal Component ───────────────────────────────────────────────────
function ConfirmModal({ visible, msg, confirmLabel, cancelLabel, onConfirm, onCancel }) {
  return (
    <Modal visible={visible} transparent animationType="fade" onRequestClose={onCancel}>
      <View style={s.confirmOverlay}>
        <View style={s.confirmBox}>
          <Text style={s.confirmMsg}>{msg}</Text>
          <View style={s.confirmBtns}>
            <TouchableOpacity onPress={onCancel} style={[s.confirmBtn, s.confirmBtnCancel]}>
              <Text style={[s.confirmBtnText, { color: C.text }]}>{cancelLabel}</Text>
            </TouchableOpacity>
            <TouchableOpacity onPress={onConfirm} style={[s.confirmBtn, s.confirmBtnDanger]}>
              <Text style={[s.confirmBtnText, { color: '#fff' }]}>{confirmLabel}</Text>
            </TouchableOpacity>
          </View>
        </View>
      </View>
    </Modal>
  );
}

// ─── Toast Component ──────────────────────────────────────────────────────────
function Toast({ msg }) {
  return (
    <View style={s.toastWrap} pointerEvents="none">
      <Text style={s.toastText}>{msg}</Text>
    </View>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
function App() {
  const insets = useSafeAreaInsets();
  const [lang, setLangState] = useState('en');
  const t = TRANSLATIONS[lang];

  const [items,       setItems]       = useState([]);
  const [customCats,  setCustomCats]  = useState([]);
  const [loaded,      setLoaded]      = useState(false);

  const categories = getAllCats(lang, customCats);

  const [view,           setView]           = useState('home');
  const [selectedId,     setSelectedId]     = useState(null);
  const [editingItem,    setEditingItem]     = useState(null);
  const [query,          setQuery]          = useState('');
  const [filterCat,      setFilterCat]      = useState(null);
  const [toast,          setToast]          = useState(null);
  const [confirmDelete,  setConfirmDelete]  = useState(null);
  const [showCatMgr,     setShowCatMgr]     = useState(false);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [newCatInput,    setNewCatInput]    = useState('');

  // ── Load data from storage on mount ──
  useEffect(() => {
    async function loadData() {
      try {
        const [storedItems, storedCats, storedLang] = await Promise.all([
          AsyncStorage.getItem(STORAGE_KEY),
          AsyncStorage.getItem(CUSTOM_CATS_KEY),
          AsyncStorage.getItem(LANG_KEY),
        ]);
        if (storedItems) setItems(JSON.parse(storedItems));
        if (storedCats)  setCustomCats(JSON.parse(storedCats));
        if (storedLang && TRANSLATIONS[storedLang]) {
          setLangState(storedLang);
        } else {
          // Auto-detect: Chinese OS → zh, everything else → en
          const locale = Localization.getLocales?.()?.[0]?.languageCode || 'en';
          if (locale.startsWith('zh')) setLangState('zh');
          else                         setLangState('en');
        }
      } catch (e) {}
      setLoaded(true);
    }
    loadData();
  }, []);

  // ── Persist on change ──
  useEffect(() => {
    if (loaded) AsyncStorage.setItem(STORAGE_KEY, JSON.stringify(items)).catch(() => {});
  }, [items, loaded]);
  useEffect(() => {
    if (loaded) AsyncStorage.setItem(CUSTOM_CATS_KEY, JSON.stringify(customCats)).catch(() => {});
  }, [customCats, loaded]);

  function setLang(l) {
    setLangState(l);
    AsyncStorage.setItem(LANG_KEY, l).catch(() => {});
  }

  function showToast(msg) {
    setToast(msg);
    setTimeout(() => setToast(null), 2200);
  }

  const filtered = items
    .filter(item => {
      const matchQ = !query
        || item.name.toLowerCase().includes(query.toLowerCase())
        || (item.location || '').toLowerCase().includes(query.toLowerCase());
      const matchC = !filterCat || item.categoryId === filterCat;
      return matchQ && matchC;
    })
    .sort((a, b) => b.createdAt - a.createdAt);

  function openDetail(id) { setSelectedId(id); setView('detail'); }
  function openForm(item = null) {
    setEditingItem(item
      ? { ...item }
      : { id: genId(), name: '', location: '', categoryId: 'other', itemPhoto: null, locationPhoto: null, createdAt: Date.now() }
    );
    setView('form');
  }
  function saveItem(item) {
    setItems(prev => {
      const exists = prev.find(i => i.id === item.id);
      return exists ? prev.map(i => i.id === item.id ? item : i) : [item, ...prev];
    });
    showToast(t.savedToast(item.name));
    setView('home');
  }
  function deleteItem(id) {
    setItems(prev => prev.filter(i => i.id !== id));
    setConfirmDelete(null);
    setView('home');
    showToast(t.deletedToast);
  }

  async function exportJson() {
    try {
      const data = JSON.stringify({ items, customCats }, null, 2);
      const filename = `cached-backup-${Date.now()}.json`;
      const fileUri = FileSystem.documentDirectory + filename;
      await FileSystem.writeAsStringAsync(fileUri, data);
      await Sharing.shareAsync(fileUri, { mimeType: 'application/json' });
      showToast(t.backupToast);
    } catch (e) {}
    setShowExportMenu(false);
  }

  async function exportXlsx() {
    try {
      const rows = items.map(item => {
        const cat = categories.find(c => c.id === item.categoryId);
        return {
          [t.colName]:     item.name,
          [t.colLocation]: item.location || '',
          [t.colCategory]: cat ? `${cat.emoji} ${cat.label}` : '',
          [t.colDate]:     new Date(item.createdAt).toLocaleDateString(t.dateLocale),
        };
      });
      const ws = XLSX.utils.json_to_sheet(rows);
      ws['!cols'] = [{ wch: 24 }, { wch: 30 }, { wch: 16 }, { wch: 16 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, t.appName);
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
      const filename = `cached-${Date.now()}.xlsx`;
      const fileUri = FileSystem.documentDirectory + filename;
      await FileSystem.writeAsStringAsync(fileUri, wbout, { encoding: FileSystem.EncodingType.Base64 });
      await Sharing.shareAsync(fileUri, { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      showToast(t.xlsxToast);
    } catch (e) {}
    setShowExportMenu(false);
  }

  const selectedItem = items.find(i => i.id === selectedId);
  const getCat = (id) => categories.find(c => c.id === id) || { label: t.defaultCats[8], emoji: '📦' };

  // ── Loading screen ──
  if (!loaded) {
    return (
      <View style={{ flex: 1, justifyContent: 'center', alignItems: 'center', backgroundColor: C.bg }}>
        <ActivityIndicator size="large" color={C.accent} />
      </View>
    );
  }

  // ── Home View ─────────────────────────────────────────────────────────────
  if (view === 'home') return (
    <View style={[s.screen, { paddingBottom: insets.bottom }]}>
      <StatusBar barStyle="dark-content" backgroundColor={C.bg} />

      {/* Header */}
      <View style={[s.header, { paddingTop: insets.top + 12 }]}>
        <View>
          <Text style={s.brandLabel}>{t.appName}</Text>
          <Text style={s.countTitle}>
            {query ? t.foundCount(filtered.length) : t.itemCount(items.length)}
          </Text>
        </View>
        <View style={s.headerRight}>
          <LangSwitcher lang={lang} setLang={setLang} />
          <TouchableOpacity onPress={() => setShowExportMenu(true)} style={[s.iconBtn, { marginLeft: 8 }]}>
            <Ionicons name="download-outline" size={20} color={C.textMuted} />
          </TouchableOpacity>
        </View>
      </View>

      {/* Search */}
      <View style={s.searchWrap}>
        <Ionicons name="search" size={18} color={C.textMuted} style={{ marginLeft: 14 }} />
        <TextInput
          value={query}
          onChangeText={setQuery}
          placeholder={t.searchPlaceholder}
          placeholderTextColor={C.textMuted}
          style={s.searchInput}
        />
        {query ? (
          <TouchableOpacity onPress={() => setQuery('')} style={{ padding: 12 }}>
            <Ionicons name="close-circle" size={18} color={C.textMuted} />
          </TouchableOpacity>
        ) : null}
      </View>

      {/* Category filter row */}
      <ScrollView horizontal showsHorizontalScrollIndicator={false} style={{ flexGrow: 0, marginBottom: 4 }}
        contentContainerStyle={{ paddingHorizontal: 20, paddingBottom: 4, gap: 8 }}>
        <TouchableOpacity onPress={() => setFilterCat(null)} style={[s.catChip, filterCat === null && s.catChipActive]}>
          <Text style={[s.catChipText, filterCat === null && s.catChipTextActive]}>{t.all}</Text>
        </TouchableOpacity>
        {categories.map(c => (
          <TouchableOpacity key={c.id} onPress={() => setFilterCat(filterCat === c.id ? null : c.id)}
            style={[s.catChip, filterCat === c.id && s.catChipActive]}>
            <Text style={[s.catChipText, filterCat === c.id && s.catChipTextActive]}>{c.emoji} {c.label}</Text>
          </TouchableOpacity>
        ))}
        <TouchableOpacity onPress={() => setShowCatMgr(true)} style={[s.catChip, { borderStyle: 'dashed' }]}>
          <Text style={s.catChipText}>{t.manage}</Text>
        </TouchableOpacity>
      </ScrollView>

      {/* Item list */}
      <FlatList
        data={filtered}
        keyExtractor={item => item.id}
        style={{ flex: 1 }}
        contentContainerStyle={{ padding: 20, paddingBottom: 100, gap: 10 }}
        showsVerticalScrollIndicator={false}
        ListEmptyComponent={() => (
          <View style={s.emptyWrap}>
            <Text style={s.emptyEmoji}>📦</Text>
            <Text style={s.emptyText}>{query ? t.emptySearch : t.emptyList}</Text>
          </View>
        )}
        renderItem={({ item }) => {
          const cat = getCat(item.categoryId);
          return (
            <TouchableOpacity onPress={() => openDetail(item.id)} style={s.itemCard}>
              <View style={s.itemThumb}>
                {item.itemPhoto
                  ? <Image source={{ uri: item.itemPhoto }} style={s.itemThumbImg} />
                  : <Text style={s.itemThumbEmoji}>{cat.emoji}</Text>
                }
              </View>
              <View style={{ flex: 1, minWidth: 0 }}>
                <Text style={s.itemName} numberOfLines={1}>{item.name}</Text>
                <View style={s.itemLocRow}>
                  <Ionicons name="location-outline" size={13} color={C.textMuted} />
                  <Text style={s.itemLoc} numberOfLines={1}>{item.location || t.noLocation}</Text>
                </View>
              </View>
              <View style={s.itemCatBadge}>
                <Text style={s.itemCatBadgeText}>{cat.emoji} {cat.label}</Text>
              </View>
            </TouchableOpacity>
          );
        }}
      />

      {/* FAB */}
      <TouchableOpacity onPress={() => openForm()} style={[s.fab, { bottom: 28 + insets.bottom }]}>
        <Ionicons name="add" size={28} color="#fff" />
      </TouchableOpacity>

      {/* Category Manager Modal */}
      <BottomModal visible={showCatMgr} onClose={() => { setShowCatMgr(false); setNewCatInput(''); }} title={t.manageCategories}>
        <View style={{ paddingBottom: 24 }}>
          <View style={s.catInputRow}>
            <TextInput
              value={newCatInput}
              onChangeText={setNewCatInput}
              placeholder={t.newCategoryPlaceholder}
              placeholderTextColor={C.textMuted}
              style={[s.input, { flex: 1, marginRight: 8 }]}
              onSubmitEditing={() => {
                if (!newCatInput.trim()) return;
                setCustomCats(prev => [...prev, { id: genId(), label: newCatInput.trim(), emoji: '🏷️' }]);
                setNewCatInput('');
              }}
            />
            <TouchableOpacity onPress={() => {
              if (!newCatInput.trim()) return;
              setCustomCats(prev => [...prev, { id: genId(), label: newCatInput.trim(), emoji: '🏷️' }]);
              setNewCatInput('');
            }} style={s.addBtn}>
              <Text style={s.addBtnText}>{t.add}</Text>
            </TouchableOpacity>
          </View>
          <View style={{ gap: 8 }}>
            {getDefaultCats(lang).map(c => (
              <View key={c.id} style={s.catMgrItem}>
                <Text style={s.catMgrItemText}>{c.emoji} {c.label}</Text>
                <Text style={{ fontSize: 12, color: C.textMuted }}>—</Text>
              </View>
            ))}
            {customCats.map(c => (
              <View key={c.id} style={s.catMgrItem}>
                <Text style={s.catMgrItemText}>{c.emoji} {c.label}</Text>
                <TouchableOpacity onPress={() => setCustomCats(prev => prev.filter(x => x.id !== c.id))}>
                  <Ionicons name="trash-outline" size={17} color={C.danger} />
                </TouchableOpacity>
              </View>
            ))}
          </View>
        </View>
      </BottomModal>

      {/* Export Modal */}
      <BottomModal visible={showExportMenu} onClose={() => setShowExportMenu(false)} title={t.exportTitle}>
        <View style={{ paddingBottom: 24, gap: 8 }}>
          <TouchableOpacity onPress={exportXlsx} style={s.exportItem}>
            <Ionicons name="grid-outline" size={18} color={C.text} />
            <Text style={s.exportItemText}>{t.exportXlsx}</Text>
          </TouchableOpacity>
          <TouchableOpacity onPress={exportJson} style={s.exportItem}>
            <Ionicons name="download-outline" size={18} color={C.text} />
            <Text style={s.exportItemText}>{t.exportJson}</Text>
          </TouchableOpacity>
        </View>
      </BottomModal>

      {toast && <Toast msg={toast} />}
    </View>
  );

  // ── Detail View ───────────────────────────────────────────────────────────
  if (view === 'detail' && selectedItem) {
    const cat = getCat(selectedItem.categoryId);
    const dateStr = new Date(selectedItem.createdAt).toLocaleDateString(t.dateLocale, {
      year: 'numeric', month: 'long', day: 'numeric',
    });
    return (
      <View style={[s.screen, { paddingBottom: insets.bottom }]}>
        <StatusBar barStyle="dark-content" backgroundColor={C.bg} />
        <View style={[s.header, { paddingTop: insets.top + 12 }]}>
          <TouchableOpacity onPress={() => setView('home')} style={s.iconBtn}>
            <Ionicons name="chevron-back" size={20} color={C.text} />
          </TouchableOpacity>
          <View style={{ flex: 1 }} />
          <LangSwitcher lang={lang} setLang={setLang} />
          <TouchableOpacity onPress={() => openForm(selectedItem)} style={[s.iconBtn, { marginLeft: 8 }]}>
            <Ionicons name="create-outline" size={18} color={C.accent} />
          </TouchableOpacity>
          <TouchableOpacity onPress={() => setConfirmDelete(selectedItem.id)} style={[s.iconBtn, { marginLeft: 8 }]}>
            <Ionicons name="trash-outline" size={18} color={C.danger} />
          </TouchableOpacity>
        </View>

        <ScrollView style={{ flex: 1 }} contentContainerStyle={{ padding: 20, paddingBottom: 40 }} showsVerticalScrollIndicator={false}>
          {selectedItem.itemPhoto ? (
            <View style={s.detailHeroWrap}>
              <Image source={{ uri: selectedItem.itemPhoto }} style={s.detailHeroImg} resizeMode="cover" />
            </View>
          ) : (
            <View style={s.detailHeroPlaceholder}>
              <Text style={{ fontSize: 52 }}>{cat.emoji}</Text>
            </View>
          )}
          <Text style={s.detailName}>{selectedItem.name}</Text>
          <View style={s.detailCatBadge}>
            <Ionicons name="pricetag-outline" size={11} color={C.accent} />
            <Text style={s.detailCatBadgeText}> {cat.emoji} {cat.label}</Text>
          </View>

          <View style={s.detailLocationCard}>
            <View style={s.detailLocationHeader}>
              <Ionicons name="location-outline" size={13} color={C.textMuted} />
              <Text style={s.detailLocationLabel}> {t.storedAt}</Text>
            </View>
            <Text style={s.detailLocationText}>
              {selectedItem.location || t.notRecorded}
            </Text>
          </View>

          {selectedItem.locationPhoto ? (
            <View style={{ marginBottom: 16 }}>
              <Text style={s.sectionLabel}>{t.locationPhotoLabel}</Text>
              <Image source={{ uri: selectedItem.locationPhoto }} style={s.detailLocPhoto} resizeMode="cover" />
            </View>
          ) : null}

          <Text style={s.detailDate}>{t.recordedOn(dateStr)}</Text>
        </ScrollView>

        <ConfirmModal
          visible={!!confirmDelete}
          msg={t.confirmDelete}
          confirmLabel={t.delete}
          cancelLabel={t.cancel}
          onConfirm={() => deleteItem(confirmDelete)}
          onCancel={() => setConfirmDelete(null)}
        />
        {toast && <Toast msg={toast} />}
      </View>
    );
  }

  // ── Form View ─────────────────────────────────────────────────────────────
  if (view === 'form' && editingItem) {
    const isNew   = !items.find(i => i.id === editingItem.id);
    const update  = (k, v) => setEditingItem(prev => ({ ...prev, [k]: v }));
    const canSave = editingItem.name.trim() && editingItem.location.trim();
    return (
      <KeyboardAvoidingView
        style={[s.screen, { paddingBottom: insets.bottom }]}
        behavior={Platform.OS === 'ios' ? 'padding' : undefined}
      >
        <StatusBar barStyle="dark-content" backgroundColor={C.bg} />
        <View style={[s.header, { paddingTop: insets.top + 12 }]}>
          <TouchableOpacity onPress={() => setView(isNew ? 'home' : 'detail')} style={s.iconBtn}>
            <Ionicons name="chevron-back" size={20} color={C.text} />
          </TouchableOpacity>
          <Text style={s.formTitle}>{isNew ? t.addItem : t.editItem}</Text>
          <View style={{ flex: 1 }} />
          <LangSwitcher lang={lang} setLang={setLang} />
          <TouchableOpacity
            onPress={() => canSave && saveItem(editingItem)}
            style={[s.saveBtn, { marginLeft: 8 }, !canSave && { opacity: 0.4 }]}
            disabled={!canSave}
          >
            <Text style={s.saveBtnText}>{t.save}</Text>
          </TouchableOpacity>
        </View>

        <ScrollView
          style={{ flex: 1 }}
          contentContainerStyle={{ padding: 20, paddingBottom: 60 }}
          keyboardShouldPersistTaps="handled"
          showsVerticalScrollIndicator={false}
        >
          {/* Category picker */}
          <Text style={s.fieldLabel}>{t.category}</Text>
          <View style={s.catPickerRow}>
            {categories.map(c => (
              <TouchableOpacity key={c.id} onPress={() => update('categoryId', c.id)}
                style={[s.catPickerChip, editingItem.categoryId === c.id && s.catPickerChipActive]}>
                <Text style={[s.catPickerChipText, editingItem.categoryId === c.id && { color: C.accent }]}>
                  {c.emoji} {c.label}
                </Text>
              </TouchableOpacity>
            ))}
          </View>

          {/* Item name */}
          <Text style={s.fieldLabel}>{t.itemName}</Text>
          <TextInput
            value={editingItem.name}
            onChangeText={v => update('name', v)}
            placeholder={t.itemNamePlaceholder}
            placeholderTextColor={C.textMuted}
            style={[s.input, { marginBottom: 8 }]}
          />
          {/* Quick-pick presets */}
          <ScrollView horizontal showsHorizontalScrollIndicator={false}
            contentContainerStyle={{ gap: 6, paddingBottom: 14 }}>
            {(PRESETS[lang]?.[editingItem.categoryId] || []).map(p => (
              <TouchableOpacity key={p} onPress={() => update('name', p)}
                style={[s.presetChip, editingItem.name === p && s.presetChipActive]}>
                <Text style={[s.presetChipText, editingItem.name === p && { color: C.accent }]}>{p}</Text>
              </TouchableOpacity>
            ))}
          </ScrollView>

          {/* Item photo */}
          <PhotoUpload
            value={editingItem.itemPhoto}
            onChange={v => update('itemPhoto', v)}
            label={t.itemPhoto}
            hint={t.itemPhotoHint}
            compressingLabel={t.compressing}
            t={t}
          />

          {/* Location */}
          <Text style={s.fieldLabel}>{t.location}</Text>
          <TextInput
            value={editingItem.location}
            onChangeText={v => update('location', v)}
            placeholder={t.locationPlaceholder}
            placeholderTextColor={C.textMuted}
            style={[s.input, { marginBottom: 16 }]}
            multiline
            numberOfLines={2}
          />

          {/* Location photo */}
          <PhotoUpload
            value={editingItem.locationPhoto}
            onChange={v => update('locationPhoto', v)}
            label={t.locationPhoto}
            hint={t.locationPhotoHint}
            compressingLabel={t.compressing}
            t={t}
          />
        </ScrollView>
      </KeyboardAvoidingView>
    );
  }

  return null;
}

// ─── Root with SafeAreaProvider ───────────────────────────────────────────────
export default App;

// ─── Styles ───────────────────────────────────────────────────────────────────
const s = StyleSheet.create({
  // Layout
  screen:              { flex: 1, backgroundColor: C.bg },
  header:              { flexDirection: 'row', alignItems: 'center', paddingHorizontal: 20, paddingBottom: 12 },
  headerRight:         { flexDirection: 'row', alignItems: 'center', marginLeft: 'auto' },
  // Header text
  brandLabel:          { fontSize: 11, letterSpacing: 2, color: C.accent, fontWeight: '700', textTransform: 'uppercase', marginBottom: 2 },
  countTitle:          { fontSize: 26, fontWeight: '800', color: C.text, letterSpacing: -0.5 },
  // Icon button
  iconBtn:             { backgroundColor: C.surface, borderWidth: 1.5, borderColor: C.border, borderRadius: 12, width: 40, height: 40, alignItems: 'center', justifyContent: 'center', flexDirection: 'row' },
  // Search
  searchWrap:          { flexDirection: 'row', alignItems: 'center', marginHorizontal: 20, marginBottom: 12, backgroundColor: C.surface, borderWidth: 1.5, borderColor: C.border, borderRadius: 14 },
  searchInput:         { flex: 1, padding: 13, paddingLeft: 10, fontSize: 16, color: C.text },
  // Category chips
  catChip:             { paddingHorizontal: 14, paddingVertical: 7, borderRadius: 20, borderWidth: 1.5, borderColor: C.border, backgroundColor: C.surface },
  catChipActive:       { borderColor: C.accent, backgroundColor: C.accentBg },
  catChipText:         { fontSize: 13, fontWeight: '600', color: C.textMuted },
  catChipTextActive:   { color: C.accent },
  // Empty state
  emptyWrap:           { alignItems: 'center', paddingTop: 60 },
  emptyEmoji:          { fontSize: 40, marginBottom: 12 },
  emptyText:           { fontSize: 15, color: C.textMuted, textAlign: 'center' },
  // Item card
  itemCard:            { flexDirection: 'row', alignItems: 'center', gap: 14, backgroundColor: C.surface, borderWidth: 1.5, borderColor: C.border, borderRadius: 16, padding: 14 },
  itemThumb:           { width: 52, height: 52, borderRadius: 12, overflow: 'hidden', backgroundColor: C.surface2, alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  itemThumbImg:        { width: '100%', height: '100%' },
  itemThumbEmoji:      { fontSize: 22 },
  itemName:            { fontSize: 16, fontWeight: '700', color: C.text, marginBottom: 3, letterSpacing: -0.3 },
  itemLocRow:          { flexDirection: 'row', alignItems: 'center', gap: 4 },
  itemLoc:             { fontSize: 13, color: C.textMuted, flex: 1 },
  itemCatBadge:        { backgroundColor: C.accentBg, paddingHorizontal: 8, paddingVertical: 3, borderRadius: 20 },
  itemCatBadgeText:    { fontSize: 11, color: C.accent, fontWeight: '600' },
  // FAB
  fab:                 { position: 'absolute', right: 24, width: 60, height: 60, borderRadius: 30, backgroundColor: C.accent, alignItems: 'center', justifyContent: 'center', shadowColor: C.accent, shadowOffset: { width: 0, height: 8 }, shadowOpacity: 0.35, shadowRadius: 12, elevation: 8 },
  // Bottom Modal
  bottomSheet:         { backgroundColor: C.bg, borderTopLeftRadius: 24, borderTopRightRadius: 24, padding: 20, paddingBottom: 10, maxHeight: '75%' },
  bottomSheetHandle:   { width: 36, height: 4, backgroundColor: C.border, borderRadius: 2, alignSelf: 'center', marginBottom: 16 },
  bottomSheetHeader:   { flexDirection: 'row', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 },
  bottomSheetTitle:    { fontSize: 18, fontWeight: '700', color: C.text },
  // Confirm Modal
  confirmOverlay:      { flex: 1, backgroundColor: 'rgba(0,0,0,0.5)', justifyContent: 'center', alignItems: 'center', padding: 24 },
  confirmBox:          { backgroundColor: C.bg, borderRadius: 20, padding: 28, width: '100%', maxWidth: 320 },
  confirmMsg:          { fontSize: 16, fontWeight: '600', color: C.text, marginBottom: 24, textAlign: 'center' },
  confirmBtns:         { flexDirection: 'row', gap: 10 },
  confirmBtn:          { flex: 1, padding: 12, borderRadius: 12, alignItems: 'center' },
  confirmBtnCancel:    { backgroundColor: C.surface, borderWidth: 1.5, borderColor: C.border },
  confirmBtnDanger:    { backgroundColor: C.danger },
  confirmBtnText:      { fontSize: 15, fontWeight: '600' },
  // Toast
  toastWrap:           { position: 'absolute', bottom: 100, left: 20, right: 20, alignItems: 'center', pointerEvents: 'none' },
  toastText:           { backgroundColor: C.text, color: C.bg, paddingHorizontal: 22, paddingVertical: 12, borderRadius: 30, fontSize: 14, fontWeight: '600', overflow: 'hidden' },
  // Lang menu
  langMenuOverlay:     { paddingTop: 80, alignItems: 'flex-end', paddingRight: 20 },
  langMenu:            { backgroundColor: C.surface, borderWidth: 1.5, borderColor: C.border, borderRadius: 14, overflow: 'hidden', minWidth: 150, shadowColor: '#000', shadowOffset: { width: 0, height: 8 }, shadowOpacity: 0.12, shadowRadius: 12, elevation: 5 },
  langItem:            { flexDirection: 'row', alignItems: 'center', gap: 10, padding: 12, paddingHorizontal: 16 },
  langItemActive:      { backgroundColor: C.accentBg },
  langItemText:        { fontSize: 14, fontWeight: '500', color: C.text },
  // Form
  formTitle:           { fontSize: 18, fontWeight: '700', color: C.text, marginLeft: 12 },
  saveBtn:             { backgroundColor: C.accent, paddingHorizontal: 20, paddingVertical: 8, borderRadius: 14 },
  saveBtnText:         { color: '#fff', fontSize: 14, fontWeight: '700' },
  fieldLabel:          { fontSize: 12, fontWeight: '600', letterSpacing: 1, color: C.textMuted, textTransform: 'uppercase', marginBottom: 8, marginTop: 8 },
  input:               { backgroundColor: C.surface, borderWidth: 1.5, borderColor: C.border, borderRadius: 12, padding: 13, fontSize: 16, color: C.text },
  catPickerRow:        { flexDirection: 'row', flexWrap: 'wrap', gap: 8, marginBottom: 16 },
  catPickerChip:       { paddingHorizontal: 14, paddingVertical: 8, borderRadius: 20, borderWidth: 1.5, borderColor: C.border, backgroundColor: C.surface },
  catPickerChipActive: { borderColor: C.accent, backgroundColor: C.accentBg },
  catPickerChipText:   { fontSize: 13, fontWeight: '600', color: C.textMuted },
  presetChip:          { paddingHorizontal: 12, paddingVertical: 5, borderRadius: 16, borderWidth: 1.5, borderColor: C.border, backgroundColor: C.surface2 },
  presetChipActive:    { backgroundColor: C.accentBg, borderColor: C.accent },
  presetChipText:      { fontSize: 12, fontWeight: '600', color: C.textMuted },
  // Photo Upload
  photoWrap:           { marginBottom: 16 },
  photoLabel:          { fontSize: 12, fontWeight: '600', letterSpacing: 1, color: C.textMuted, textTransform: 'uppercase', marginBottom: 8, marginTop: 8 },
  photoPlaceholder:    { width: '100%', height: 110, borderWidth: 2, borderStyle: 'dashed', borderColor: C.border, borderRadius: 12, backgroundColor: C.surface2, alignItems: 'center', justifyContent: 'center' },
  photoHint:           { fontSize: 13, color: C.textMuted },
  photoPreviewWrap:    { borderRadius: 12, overflow: 'hidden' },
  photoPreview:        { width: '100%', height: 180 },
  photoRemoveBtn:      { position: 'absolute', top: 8, right: 8, backgroundColor: 'rgba(0,0,0,0.6)', borderRadius: 15, width: 30, height: 30, alignItems: 'center', justifyContent: 'center' },
  // Cat manager
  catInputRow:         { flexDirection: 'row', alignItems: 'center', marginBottom: 16 },
  addBtn:              { backgroundColor: C.accent, paddingHorizontal: 16, paddingVertical: 13, borderRadius: 12 },
  addBtnText:          { color: '#fff', fontSize: 14, fontWeight: '700' },
  catMgrItem:          { flexDirection: 'row', justifyContent: 'space-between', alignItems: 'center', padding: 10, paddingHorizontal: 14, backgroundColor: C.surface2, borderRadius: 10, marginBottom: 8 },
  catMgrItemText:      { fontSize: 15, color: C.text },
  // Export
  exportItem:          { flexDirection: 'row', alignItems: 'center', gap: 10, padding: 13, paddingHorizontal: 16, backgroundColor: C.surface2, borderRadius: 12, marginBottom: 8 },
  exportItemText:      { fontSize: 14, fontWeight: '500', color: C.text },
  // Detail
  detailHeroWrap:      { borderRadius: 18, overflow: 'hidden', marginBottom: 20, aspectRatio: 16 / 9 },
  detailHeroImg:       { width: '100%', height: '100%' },
  detailHeroPlaceholder: { borderRadius: 18, backgroundColor: C.surface2, marginBottom: 20, height: 120, alignItems: 'center', justifyContent: 'center' },
  detailName:          { fontSize: 28, fontWeight: '800', color: C.text, letterSpacing: -0.5, marginBottom: 6 },
  detailCatBadge:      { flexDirection: 'row', alignItems: 'center', backgroundColor: C.accentBg, paddingHorizontal: 10, paddingVertical: 4, borderRadius: 20, alignSelf: 'flex-start', marginBottom: 24 },
  detailCatBadgeText:  { fontSize: 12, color: C.accent, fontWeight: '700' },
  detailLocationCard:  { backgroundColor: C.surface, borderWidth: 1.5, borderColor: C.border, borderRadius: 16, padding: 18, marginBottom: 16 },
  detailLocationHeader:{ flexDirection: 'row', alignItems: 'center', marginBottom: 10 },
  detailLocationLabel: { fontSize: 11, fontWeight: '700', letterSpacing: 1, color: C.textMuted, textTransform: 'uppercase' },
  detailLocationText:  { fontSize: 17, color: C.text, fontWeight: '600', lineHeight: 26 },
  sectionLabel:        { fontSize: 11, fontWeight: '700', letterSpacing: 1, color: C.textMuted, textTransform: 'uppercase', marginBottom: 8 },
  detailLocPhoto:      { width: '100%', borderRadius: 14 },
  detailDate:          { fontSize: 12, color: C.textMuted, marginTop: 8 },
});
