// Because Excel ClientObjects all implement their own load functions, we define
// a generic interface here for use with fetch()
export interface LoadableClientObject extends OfficeExtension.ClientObject {
  load(options?: Record<string, any>): OfficeExtension.ClientObject
  load(propertyNames?: string | string[]): OfficeExtension.ClientObject
  load(propertyNamesAndPaths?: { select?: string, expand?: string }): OfficeExtension.ClientObject
}

export type FetchBatch = Record<string, LoadableClientObject>

export default class Context {
  context: Excel.RequestContext

  constructor(context: Excel.RequestContext) {
    this.context = context
  }

  addTrackedObject(object: OfficeExtension.ClientObject) {
    this.context.trackedObjects.add(object)
  }

  removeTrackedObject(object: OfficeExtension.ClientObject) {
    this.context.trackedObjects.remove(object)
  }

  async fetch<T extends FetchBatch>(
    createBatch: (ctx: Excel.RequestContext) => Promise<T>
  ): Promise<T> {
    const batch = await createBatch(this.context)

    // TODO: don't load if already loaded
    Object.keys(batch).forEach((key) => {
      batch[key].load()
    })

    await this.context.sync()

    return batch
  }

  async sync(operation: (ctx: Excel.RequestContext) => Promise<void>): Promise<void> {
    await operation(this.context)
    await this.context.sync()
  }
}
