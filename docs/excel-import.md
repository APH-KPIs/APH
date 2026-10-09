# قالب Excel الرسمي ونظام الاستيراد

القالب `templates/APH_KPI_Data_Template_2026.xlsx` هو قالب الإدخال الرسمي والمصدر الرئيسي لبيانات المنصة. يبدأ التعبئة من **01/01/2026** ولا يوجد تاريخ نهاية.

مسار البيانات:

```
Excel  →  Validation (في المتصفح)  →  Mapping Layer  →  Database (فحص ثانٍ في الخادم)  →  KPI Engine  →  Dashboard
```

الملف يُعامل كمدخل **غير موثوق (UNTRUSTED INPUT)**: حماية Excel وقوائمه تمنع الأخطاء الشائعة فقط، والفحص الحقيقي يتم مرتين، في صفحة «إدارة البيانات» ثم في قاعدة البيانات (`import_rows` في `supabase/setup.sql`). لا يعتمد الاستيراد على الألوان أو الخلايا المدمجة أو ترتيب الأعمدة، بل على أسماء الأعمدة الثابتة.

## القالب البسيط (للتعبئة اليومية)

`templates/APH_KPI_Simple_2026.xlsx` (يُحمّل من المنصة باسم «قالب إدخال المؤشرات 2026») هو الملف الذي يُعبأ عادة. كل نصوصه عربية:

- ورقة لكل شهر (**يناير** إلى **ديسمبر**)، أفقية: المؤشرات في الصفوف (صفّان للبسط والمقام في مؤشرات النسب) وأيام الشهر في الأعمدة، من 01/01/2026 إلى 31/12/2026. الخلايا الرمادية ليست يوم إدخال المؤشر (الأسبوعي يوم الأحد، والشهري أول الشهر، والربعي أول الربع).
- ورقة **المؤشرات**: الإدارة والدورية ويوم الإدخال والوحدة والمستهدف والتعريف.
- ورقة **التعليمات**.

تتعرف المنصة على الصفوف من عمود رموز مخفي، أو من اسم المؤشر العربي إذا حُذف ذلك العمود (ويُقبل أيضاً الشكل العمودي: سطر لكل يوم وعمود لكل مؤشر)، وتحوّل كل خلية معبأة إلى سجل للإدارة المسؤولة عن المؤشر والمنشأة الرئيسية، ثم تمر بالفحص نفسه المذكور أدناه. يُبنى الملف بـ `python3 tools/build_simple_template.py`.

القالب التفصيلي التالي (`APH_KPI_Data_Template_2026.xlsx`) يُستخدم عند الحاجة لعدة منشآت أو إدارات، ويستورد بالطريقة نفسها.

---

## 1. أوراق القالب

| الورقة | الغرض | من يعدّلها |
|---|---|---|
| `README` | اسم الملف والغرض وطريقة التعبئة والحقول والصيغ والوحدات وطريقة الرفع ورموز الأخطاء، مع مثال صحيح ومثال خاطئ. يحمل رقم الإصدار `Template_Version`. | لا أحد |
| `DATA` | إدخال البيانات: سجل واحد لكل سطر (Long Format). | مدخل البيانات (الخلايا الصفراء فقط) |
| `INDICATORS` | جدول المؤشرات الرئيسي (Master). | مدير النظام |
| `DEPARTMENTS` | الإدارات والأقسام. | مدير النظام |
| `FACILITIES` | المنشآت. | مدير النظام |
| `LOOKUPS` | القوائم: أنواع الفترات، الاتجاهات، أنواع الحساب، الوحدات، الحالات، التصنيفات، نعم/لا. | لا أحد |
| `VALIDATION` | نتيجة فحص كل سطر (VALID / INVALID / WARNING) ورموز المشكلات والحل المقترح وملخص بالأعداد. | محسوبة |
| `IMPORT_TEMPLATE` | 13 عموداً بترتيب ثابت، للصق بيانات جاهزة من نظام آخر. تُترك فارغة عند التعبئة في `DATA`، وعندها تقرأ المنصة `DATA`. | اختياري |

الأوراق محمية بكلمة مرور (`APH-KPI-2026`) لمنع التعديل غير المقصود فقط، وليست وسيلة أمان.

### التقويم المعبأ مسبقاً
ورقة `DATA` معبأة بتواريخ 2026 كاملة **حسب دورية كل مؤشر** وليس يومياً لكل المؤشرات: المؤشر اليومي له 365 سطراً، والأسبوعي سطر لكل أحد، والشهري سطر لكل شهر، وهكذا (7,223 سطراً). بعدها 1,500 سطر فارغ جاهز بالصيغ لأي سجل إضافي (منشأة أو إدارة أخرى أو تصحيح). السطر الذي لا يحمل قيمة يُتجاهل عند الاستيراد.

---

## 2. أعمدة ورقة DATA

الصف 1 عناوين عربية للقراءة، والصف 2 **أسماء الأعمدة الثابتة** التي يقرؤها النظام، والبيانات من الصف 3. لا تغيّر الصف 2.

| العمود | النوع | الإلزام | المصدر |
|---|---|---|---|
| Record_ID | نص | تلقائي | رقم تسلسلي |
| Unique_Key | نص | تلقائي | صيغة: `FAC-IND-YYYYMMDD-DEP-PERIOD` |
| Facility_ID | نص (قائمة) | **إلزامي** | إدخال |
| Facility_Name | نص | تلقائي | من FACILITIES |
| Department_ID | نص (قائمة) | **إلزامي** | إدخال |
| Department_Name | نص | تلقائي | من DEPARTMENTS |
| Indicator_ID | نص (قائمة) | **إلزامي** | إدخال |
| Indicator_Name، Category_ID، Category_Name | نص | تلقائي | من INDICATORS |
| Period_Type | DAILY / WEEKLY / MONTHLY / QUARTERLY / ANNUAL | **إلزامي** | إدخال، ويجب أن يطابق دورية المؤشر |
| Measurement_Date | تاريخ `DD/MM/YYYY` | **إلزامي** | إدخال، من 01/01/2026 حتى اليوم |
| Period_Start، Period_End | تاريخ | تلقائي | الأسبوع يبدأ الأحد |
| Numerator | رقم ≥ 0 | إلزامي لمؤشرات النسب | إدخال |
| Denominator | رقم > 0 | إلزامي لمؤشرات النسب | إدخال |
| Actual_Value | رقم | إلزامي لغير النسب | إدخال |
| Calculated_Value | رقم | تلقائي | البسط ÷ المقام × المضاعف، أو القيمة الفعلية |
| Target_Value | رقم | تلقائي | من INDICATORS (للعرض فقط) |
| Unit | نص | تلقائي | من INDICATORS |
| Achievement_Percentage | رقم | تلقائي | **لا يُكتب يدوياً** |
| Status | نص | تلقائي | محقق / تحذير / غير محقق / بلا مستهدف / معلوماتي / غير متاح (N/A) |
| Data_Source، Data_Entry_By، Notes | نص | اختياري | إدخال (الملاحظات حتى 1000 حرف) |
| Data_Entry_Date، Last_Updated | تاريخ | اختياري | إدخال |
| Validation_Status | VALID / INVALID / WARNING | تلقائي | من VALIDATION |

الأعمدة المخفية التي تبدأ بـ `_` أعمدة مساعدة للصيغ.

### الحسابات
| Calculation_Type | القيمة | التجميع في المنصة |
|---|---|---|
| COUNT | القيمة الفعلية | مجموع الفترة |
| SUM | القيمة الفعلية | مجموع الفترة |
| AVERAGE | القيمة الفعلية | متوسط القراءات |
| WEIGHTED_AVERAGE | القيمة الفعلية | متوسط موزون |
| LAST_VALUE | القيمة الفعلية | آخر قيمة |
| TARGET_ACHIEVEMENT | القيمة الفعلية | متوسط |
| PERCENTAGE، COMPLIANCE، RATE | البسط ÷ المقام × 100 | Σبسط ÷ Σمقام × 100 |
| RATIO | البسط ÷ المقام | Σبسط ÷ Σمقام |
| RATE_PER_1000 | البسط ÷ المقام × 1,000 | Σبسط ÷ Σمقام × 1,000 |
| RATE_PER_100000 | البسط ÷ المقام × 100,000 | Σبسط ÷ Σمقام × 100,000 |

مؤشرات النسب تقبل أيضاً النسبة برقمها النهائي بلا بسط ومقام (كما في ملف الإدخال المعتمد 2026): إذا كان للمؤشر صف واحد في ورقة الشهر، تُعتمد قيمته كما هي مهما كان رمزه (`:NUM` أو `:DEN`). يُسجَّل الربع حسب التاريخ الذي أُدخلت فيه القراءة.

نسبة الإنجاز حسب الاتجاه: `HIGHER_IS_BETTER` = الفعلي ÷ المستهدف × 100؛ `LOWER_IS_BETTER` = 100 إذا كان الفعلي ضمن المستهدف وإلا المستهدف ÷ الفعلي × 100؛ `TARGET_BASED` = القرب من المستهدف؛ `INFORMATIONAL` بلا تقييم. أي قسمة على صفر أو قيمة ناقصة تظهر **N/A**، ولا يظهر `#DIV/0!` أو `NaN` أو `Infinity`.

مؤشرات العدد والمجموع تُقيَّم في المنصة على مستوى الفترة (مستهدف الربع)، لذلك تظهر حالة السطر الواحد «يُقيَّم على مستوى الفترة».

---

## 3. أعمدة INDICATORS و DEPARTMENTS و FACILITIES

**INDICATORS:** Indicator_ID، Indicator_Name، Arabic_Name، English_Name، Category_ID، Category، Sub_Category، Department_ID، Department، Frequency، Unit، Calculation_Type، Direction، Target، Minimum_Acceptable، Maximum_Acceptable، Numerator_Definition، Denominator_Definition، Data_Source، Active، Effective_From، Effective_To، Input_Mode (INPUT/DERIVED)، Green_Threshold، Yellow_Threshold، Target_Basis، Definition.

**DEPARTMENTS:** Department_ID، Department_Name، Department_Name_AR، Parent_Department_ID، Department_Type، Active.

**FACILITIES:** Facility_ID، Facility_Name، Facility_Name_AR، Cluster، Region، Active، Is_Primary.

المستهدفات تُدار من صفحة المستهدفات في المنصة (بمبرر وسجل تدقيق). عمود Target في القالب للعرض فقط، ويُحفظ مع كل سجل مستورد في `target_value` كلقطة للمستهدف وقت الاستيراد.

إضافة مؤشر أو إدارة أو منشأة جديدة **لا تحتاج تعديل قاعدة البيانات**: أضف السطر في ورقته، وعند الرفع تعرض المنصة إضافته (لمدير النظام)، أو أضفه من المنصة مباشرة.

---

## 4. المفتاح الفريد وقواعد التكرار

```
Unique_Key = Facility_ID - Indicator_ID - YYYYMMDD(بداية الفترة) - Department_ID - Period_Type
مثال: FAC001-ED_VISITS-20260315-DEP005-DAILY
```

- يُحسب من **بداية الفترة**، لذا أي تاريخ داخل الأسبوع/الشهر/الربع يُسجل على بدايته (تنبيه W02).
- في قاعدة البيانات عليه قيد فريد (`records_record_key_uidx`).
- **إعادة رفع نفس المفتاح تحدّث السجل (UPDATE) ولا تكرره.** المطابق تماماً لا يتغير (Unchanged).
- إذا تكرر المفتاح داخل الملف نفسه يُعتمد **آخر سطر**، وتنبه شاشة الفحص في المنصة على الأسطر السابقة (W05).
- السجلات القديمة (قبل هذا الإصدار) يُحسب لها المفتاح تلقائياً عند تشغيل `setup.sql`.

---

## 5. رموز الفحص

الرموز نفسها في ورقة VALIDATION وشاشة الاستيراد وتقرير الأخطاء وجدول `import_errors`.

| الرمز | المستوى | المعنى |
|---|---|---|
| E01 | ERROR | قيمة القياس ناقصة |
| E02 | ERROR | تاريخ غير صالح |
| E03 | ERROR | تاريخ قبل 01/01/2026 أو مستقبلي (التاريخ المستقبلي تمنعه قائمة Excel وتفحصه المنصة) |
| E04 | ERROR | مؤشر غير موجود |
| E05 | ERROR | إدارة غير موجودة أو غير نشطة |
| E06 | ERROR | منشأة غير موجودة أو غير نشطة |
| E07 | ERROR | نص في حقل رقمي |
| E08 | ERROR | المقام صفر |
| E09 | ERROR | قيمة سالبة أو خارج الحد (مثل نسبة > 100%) |
| E11 | ERROR | سجل بدون مفتاح فريد |
| E12 | ERROR | نوع الفترة لا يطابق دورية المؤشر |
| E13 | ERROR | مؤشر غير نشط |
| E14 | ERROR | مؤشر محسوب تلقائياً لا يُدخل |
| E16 | ERROR | سجل معتمد أو لمستخدم آخر (المنصة فقط) |
| E17 | ERROR | رفض الخادم السجل (المنصة فقط) |
| W01 | WARNING | لا يوجد مستهدف |
| W02 | WARNING | التاريخ سُجل على بداية الفترة |
| W03 | WARNING | البسط أكبر من المقام |
| W06 | WARNING | نسبة مدخلة برقمها النهائي وتتجاوز 100%، تُستورد كما هي |
| W05 | DUPLICATE | مكرر في الملف، يُعتمد آخر سطر (المنصة فقط) |

السطر الذي فيه ERROR يُرفض وحده، وبقية الأسطر تُستورد. السطر الذي فيه WARNING يُستورد مع التنبيه.

---

## 6. خطوات الاستيراد (صفحة «إدارة البيانات»)

1. **STEP 1** قراءة الملف: الصيغة (.xlsx) والحجم (حتى 30 MB).
2. **STEP 2** إيجاد ورقة الاستيراد: `IMPORT_TEMPLATE` إذا عُبئت، وإلا ورقة `DATA`.
3. **STEP 3** إيجاد الأعمدة **بالاسم** (Mapping Layer)، فلا يهم ترتيبها، ويُقبل الاسم العربي أو الإنجليزي.
4. **STEP 4** فحص الأنواع: أرقام وتواريخ صارمة.
5. **STEP 5** فحص البيانات المرجعية: المؤشر والإدارة والمنشأة ونوع الفترة.
6. **STEP 6** فحص القيم: الحدود والمقام والبسط.
7. **STEP 7** المفتاح الفريد والتكرار داخل الملف.
8. **STEP 8** المقارنة بالمحفوظ: جديد / تحديث / مطابق.
9. **STEP 9** شاشة **Preview**: Records Found، New Records، Records to Update، Unchanged، Duplicates، Errors، Warnings، Empty Rows، وجدول الأخطاء (Row، Field، Value، Error، Suggested Fix) مع زر **Download Error Report (Excel)**.

ثم زر **VALIDATE & IMPORT** / **IMPORT & SAVE**، فترسل المنصة السجلات على دفعات من 500 إلى `import_start` ← `import_rows` ← `import_finish`، والخادم يعيد فحص كل سطر. تظهر شاشة **Import Successful** بالأعداد ورقم الدفعة، وتتحدث اللوحة عند جميع المستخدمين.

### تبويبات الصفحة
- **رفع ملف:** Upload و Download Template و Validate.
- **سجل الاستيراد (Import History):** Batch، File، Date، User، Records، Inserted، Updated، Errors، Status، مع تفاصيل كل دفعة وتقرير أخطائها.
- **جودة البيانات:** Total، Valid، Invalid، Missing Values، Duplicates، Outliers، Missing Targets، Inactive Indicators، Last Import، Last Update، Data Completeness %، Data Accuracy %.
- **سجل التعديلات (Audit Log):** القيمة القديمة والجديدة والمستخدم والوقت والسبب، مع التصفية والتصدير.
- **البيانات المرجعية:** المنشآت والإدارات والتصنيفات.

---

## 7. قاعدة البيانات

| الجدول | الغرض |
|---|---|
| `kpis` | المؤشرات (كما كانت) |
| `facilities`، `departments`، `categories` | البيانات المرجعية |
| `records` | القراءات، مع الأعمدة الجديدة: `record_key`، `facility_id`، `department_id`، `period_type`، `period_end`، `target_value`، `unit`، `data_entry_by`، `data_entry_date`، `batch_id`، `source_row` |
| `import_batches` | دفعات الاستيراد وأعدادها وحالتها |
| `import_errors` | أخطاء كل دفعة (سطر، حقل، قيمة، رمز، رسالة، حل مقترح). للإضافة فقط |
| `record_history` | سجل كل تغيير: INSERT / UPDATE / APPROVE / SOFT_DELETE / RESTORE، بالقيم القديمة والجديدة والمستخدم والوقت والدفعة والسبب. للإضافة فقط |

العرضان `indicators` و `indicator_data` يقدمان المؤشرات والقراءات بأسماء الأعمدة الموحدة.

- **الحذف ناعم فقط** (`deleted = true`)؛ الحذف الفعلي من `records` ممنوع بقاعدة في قاعدة البيانات.
- سجل التعديلات يُكتب بمشغّل (trigger) في قاعدة البيانات، فيشمل الاستيراد والإدخال اليدوي والاعتماد.
- كل الكتابة عبر دوال تتحقق من الجلسة والصلاحية.

### الربط بين Excel وقاعدة البيانات

| عمود Excel | الحقل في المنصة | عمود قاعدة البيانات |
|---|---|---|
| Unique_Key | recordKey | `records.record_key` (يُعاد حسابه في الخادم) |
| Facility_ID | facilityId | `records.facility_id` |
| Department_ID | departmentId | `records.department_id` |
| Indicator_ID | kpiCode | `records.kpi_code` |
| Period_Type | periodType | `records.period_type` |
| Measurement_Date | date | `records.date` (بداية الفترة) و `period_end` و `year` و `quarter` |
| Numerator | numerator | `records.numerator` |
| Denominator | denominator | `records.denominator` |
| Actual_Value / Calculated_Value | actualValue | `records.actual` (يحسبه الخادم للنسب) |
| Target_Value | — | `records.target_value` (لقطة من المنصة) |
| Unit | — | `records.unit` |
| Data_Source | dataSource | `records.data_source` |
| Data_Entry_By | dataEntryBy | `records.data_entry_by` |
| Data_Entry_Date | dataEntryDate | `records.data_entry_date` |
| Notes | notes | `records.notes` |
| رقم السطر | rowNumber | `records.source_row` و `import_errors.row_number` |

---

## 8. مصدر واحد للبيانات المرجعية

القوائم المرجعية وأنواع الحساب ورموز الفحص معرّفة في `index.html` (`MASTER_FACILITIES`، `MASTER_DEPARTMENTS`، `MASTER_CATEGORIES`، `CALC_TYPES`، `IMPORT_CODES`، `KPI_TEMPLATE_META`). المولّد يقرؤها عبر `tools/extract_kpis.js`، فيبني القالب ويكتب بذرة البيانات المرجعية في `setup.sql` بين `-- BEGIN MASTER DATA` و `-- END MASTER DATA`.

لإعادة بناء القالب بعد أي تغيير:

```
pip install openpyxl
python3 tools/build_excel_template.py            # القالب + بذرة setup.sql
python3 tools/build_excel_template.py --year 2027 --out templates/APH_KPI_Data_Template_2027.xlsx
```

## 9. التحديث على المنصة الحالية
1. شغّل `supabase/setup.sql` كاملاً من SQL Editor في Supabase (آمن للتكرار، ويحافظ على البيانات الحالية ويحسب مفاتيحها).
2. يدخل مدير النظام مرة واحدة ليحفظ تحديث المؤشرات (أنواع الحساب والأسماء الإنجليزية والحدود).
