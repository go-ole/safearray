# vx.x.x

**Still in development**

# Changes from go-ole

 * No `SafeArrayConversion` object. These have been moved to `Array` object.
 * `SafeArray` object has been renamed to `COMArray`.
 * All `SafeArray*()` functions are public and have been renamed.

	"SafeArray" has been removed and the names have been expanded to play better with Go naming standards.
 * CGO is supported. (**Still in development. Not tested**)
 * All known types are supported and will append to existing type.

	This is done using reflection. The purpose is to allow for any type of Go slice or array to work with SafeArrays and not have to worry too much about creating specific functions for each type to add support.

	The only difference, is that strings and other types may require manual cleanup, but user defined types do not have this problem.
 * `UnmarshalArray()` exists as a single point of creating a COM SafeArray object.
 * `MarshalArray()` exists as a single point to convert COM SafeArray object to Go slice.
 * Multidimensional COM SafeArray are supported.

# Features

 * All SafeArray functions are available and implemented. (IRecordInfo is not available by default)
 * Conversion for Byte array and String arrays exist.
 * `Array` object provides helper methods for all available SafeArray functions.
 * `Array` object provides method for retrieving total number of elements in all dimensions.
 * `Array` object provides method for appending SafeArray elements to existing Go slice of any type.
 * `Array` object provides automatically returning Go slice based on SafeArray variant type.
 * Provides helper for creating COM SafeArray from any Go slice.

**In Progress**
 * Fully documented.
 * Fully tested.
 * Supports Go multidimensional slices.
