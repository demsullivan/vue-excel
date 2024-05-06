# Simple Routing

## Without Custom Components

```vue
<template>
  <Workbook>
    <Worksheet name="Sheet1" @changed="sheetChanged">
      <h1>Hello world!</h1>
    </Worksheet>
  </Workbook>
</template>
```

## With Custom Components

### App.vue
```vue
<template>
  <Workbook>
    <MyComponent />
  </Workbook>
</template>
```

### MyComponent.vue
```vue
<template>
  <Worksheet name="Sheet1" @changed="sheetChanged">
    <h1>Hello world!</h1>
  </Worksheet>
</template>
```
