export class PackageException extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'PackageException';
    Object.setPrototypeOf(this, new.target.prototype);
  }
}
