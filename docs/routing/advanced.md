# Advanced Routing

```typescript
import { connectExcel } from 'vue-excel'
import FirstView from './FirstView.vue'
import SecondView from './SecondView.vue'

const routes = [
  { sheetName: "Sheet1", component: FirstView },
  { namedRef: "name", value: "SecondView", component: SecondView, props: ['thing', 'otherThing'] },
  { 
    async activated(ctx: Context, worksheet: Excel.Worksheet) {
      return worksheet.name == "Sheet1"
    },
    component: FirstView
  }
]

const vueExcel = await connectExcel(routes)
```