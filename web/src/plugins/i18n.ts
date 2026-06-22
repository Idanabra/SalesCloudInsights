import { createI18n } from 'vue-i18n'

export default createI18n({
  legacy: false,
  locale: localStorage.getItem('outlook-addin-lang') || 'he',
  fallbackLocale: 'en',
  messages: {
    en: {
      title:         'Opportunities / Projects',
      search:        'Search opportunities...',
      save:          'Save Email to Opportunity / Project',
      settings:      'Settings',
      refresh:       'Refresh',
      loading:       'Loading opportunities...',
      noOpps:        'No opportunities found.',
      savingToSap:   'Saving to SAP...',
      savedSuccess:  'Email saved to SAP',
      saveFailed:    'Save failed',
      salesCycleId:  'Sales Cycle ID (optional)',
      salesCycleHint:'Leave empty to show all opportunities',
      cancel:        'Cancel',
      saveSettings:  'Save & Connect',
      owner:         'Owner',
      id:            'ID',
    },
    he: {
      title:         'הזדמנויות / פרויקטים',
      search:        'חיפוש הזדמנויות...',
      save:          'שמור מייל בהזדמנות / פרויקט',
      settings:      'הגדרות',
      refresh:       'רענן',
      loading:       'טוען הזדמנויות...',
      noOpps:        'לא נמצאו הזדמנויות.',
      savingToSap:   'שומר ב-SAP...',
      savedSuccess:  'המייל נשמר ב-SAP',
      saveFailed:    'שמירה נכשלה',
      salesCycleId:  'Sales Cycle ID (אופציונלי)',
      salesCycleHint:'השאר ריק להצגת כל ההזדמנויות',
      cancel:        'ביטול',
      saveSettings:  'שמור והתחבר',
      owner:         'בעלים',
      id:            'מזהה',
    }
  }
})
