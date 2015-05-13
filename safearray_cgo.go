// #include "OleAuto.h"
// +build windows,cgo
// XXX: This needs to be tested and completed.
// This is incomplete and will not work. Really just a skeleton.

package com

import "C"

import "unsafe"
import "github.com/go-ole/com"

// AccessData returns raw array.
//
// AKA: SafeArrayAccessData in Windows API.
func AccessData(safearray *COMArray) (element uintptr, err error) {
	err = com.MaybeError(C.SafeArrayAccessData(safearray, unsafe.Pointer(&element)))
	return
}

// UnaccessData releases raw array.
//
// AKA: SafeArrayUnaccessData in Windows API.
func UnaccessData(safearray *COMArray) error {
	return com.MaybeError(C.SafeArrayUnaccessData(safearray))
}

// AllocateArrayData allocates SafeArray.
//
// AKA: SafeArrayAllocData in Windows API.
func AllocateArrayData(safearray *COMArray) error {
	return com.MaybeError(C.SafeArrayAllocData(safearray))
}

// AllocateArrayDescriptor allocates SafeArray.
//
// AKA: SafeArrayAllocDescriptor in Windows API.
func AllocateArrayDescriptor(dimensions uint32) (safearray *COMArray, err error) {
	err = com.MaybeError(C.SafeArrayAllocDescriptor(dimensions, unsafe.Pointer(&safearray)))
	return
}

// AllocateArrayDescriptorEx allocates SafeArray.
//
// AKA: SafeArrayAllocDescriptorEx in Windows API.
func AllocateArrayDescriptorEx(variantType com.VariantType, dimensions uint32) (safearray *COMArray, err error) {
	err = com.MaybeError(C.SafeArrayAllocDescriptorEx(uint16(variantType), dimensions, unsafe.Pointer(&safearray)))
	return
}

// Duplicate returns copy of SafeArray.
//
// AKA: SafeArrayCopy in Windows API.
func Duplicate(original *COMArray) (safearray *COMArray, err error) {
	err = com.MaybeError(C.SafeArrayCopy(original, unsafe.Pointer(&safearray)))
	return
}

// DuplicateData duplicates SafeArray into another SafeArray object.
//
// AKA: SafeArrayCopyData in Windows API.
func DuplicateData(original, duplicate *COMArray) error {
	return com.MaybeError(C.SafeArrayCopyData(original, unsafe.Pointer(&duplicate)))
}

// CreateArray creates SafeArray.
//
// AKA: SafeArrayCreate in Windows API.
func CreateArray(variantType com.VariantType, dimensions uint32, bounds *Bounds) (safearray *COMArray, err error) {
	sa, _, err := C.SafeArrayCreate(uint16(variantType), dimensions, bounds)
	safearray = (*COMArray)(unsafe.Pointer(&sa))
	return
}

// CreateArrayEx creates SafeArray.
//
// AKA: SafeArrayCreateEx in Windows API.
func CreateArrayEx(variantType com.VariantType, dimensions uint32, bounds *Bounds, extra uintptr) (safearray *COMArray, err error) {
	sa, _, err := C.SafeArrayCreateEx(uint16(variantType), dimensions, bounds, extra)
	safearray = (*COMArray)(unsafe.Pointer(sa))
	return
}

// CreateArrayVector creates SafeArray.
//
// AKA: SafeArrayCreateVector in Windows API.
func CreateArrayVector(variantType com.VariantType, lowerBound int32, length uint32) (safearray *COMArray, err error) {
	sa, _, err := C.SafeArrayCreateVector(uint16(variantType), lowerBound, length)
	safearray = (*COMArray)(unsafe.Pointer(sa))
	return
}

// CreateArrayVectorEx creates SafeArray.
//
// AKA: SafeArrayCreateVectorEx in Windows API.
func CreateArrayVectorEx(variantType com.VariantType, lowerBound int32, length uint32, extra uintptr) (safearray *COMArray, err error) {
	sa, _, err := C.SafeArrayCreateVectorEx(uint16(variantType), lowerBound, length, extra)
	safearray = (*COMArray)(unsafe.Pointer(sa))
	return
}

// Destroy destroys SafeArray object.
//
// AKA: SafeArrayDestroy in Windows API.
func Destroy(safearray *COMArray) error {
	return com.MaybeError(C.SafeArrayDestroy(safearray))
}

// DestroyData destroys SafeArray object.
//
// AKA: SafeArrayDestroyData in Windows API.
func DestroyData(safearray *COMArray) error {
	return com.MaybeError(C.SafeArrayDestroyData(safearray))
}

// DestroyDescriptor destroys SafeArray object.
//
// AKA: SafeArrayDestroyDescriptor in Windows API.
func DestroyDescriptor(safearray *COMArray) error {
	return com.MaybeError(C.SafeArrayDestroyDescriptor(safearray))
}

// GetDimensions is the amount of dimensions in the SafeArray.
//
// SafeArrays may have multiple dimensions. Meaning, it could be
// multidimensional array.
//
// AKA: SafeArrayGetDim in Windows API.
func GetDimensions(safearray *COMArray) (dimensions uint32, err error) {
	l, _, err := C.SafeArrayGetDim(safearray)
	dimensions = *(*uint32)(unsafe.Pointer(l))
	return
}

// GetElementSize is the element size in bytes.
//
// AKA: SafeArrayGetElemsize in Windows API.
func GetElementSize(safearray *COMArray) (length uint32, err error) {
	l, _, err := C.SafeArrayGetElemsize(safearray)
	length = *(*uint32)(unsafe.Pointer(l))
	return
}

// GetElement retrieves element at given index.
func GetElement(safearray *COMArray, index int64) (element interface{}, err error) {
	err = GetElementDirect(safearray, index, &element)
	return
}

// GetElementDirect retrieves element value at given index.
//
// AKA: SafeArrayGetElement in Windows API.
func GetElementDirect(safearray *COMArray, index int64, element interface{}) error {
	return com.MaybeError(C.SafeArrayGetElement(safearray, index, unsafe.Pointer(&element)))
}

// GetInterfaceID is the InterfaceID of the elements in the SafeArray.
//
// AKA: SafeArrayGetIID in Windows API.
func GetInterfaceID(safearray *COMArray) (guid *com.GUID, err error) {
	err = com.MaybeError(C.SafeArrayGetIID(safearray, unsafe.Pointer(&guid)))
	return
}

// GetLowerBound returns lower bounds of SafeArray.
//
// SafeArrays may have multiple dimensions. Meaning, it could be
// multidimensional array.
//
// AKA: SafeArrayGetLBound in Windows API.
func GetLowerBound(safearray *COMArray, dimension uint32) (lowerBound int64, err error) {
	err = com.MaybeError(C.SafeArrayGetLBound(safearray, dimension, unsafe.Pointer(&lowerBound)))
	return
}

// GetUpperBound returns upper bounds of SafeArray.
//
// SafeArrays may have multiple dimensions. Meaning, it could be
// multidimensional array.
//
// AKA: SafeArrayGetUBound in Windows API.
func GetUpperBound(safearray *COMArray, dimension uint32) (upperBound int64, err error) {
	err = com.MaybeError(C.SafeArrayGetUBound(safearray, dimension, unsafe.Pointer(&upperBound)))
	return
}

// GetVariantType returns data type of SafeArray.
//
// AKA: SafeArrayGetVartype in Windows API.
func GetVariantType(safearray *COMArray) (varType uint16, err error) {
	err = com.MaybeError(C.SafeArrayGetVartype(safearray, unsafe.Pointer(&varType)))
	return
}

// Lock locks SafeArray for reading to modify SafeArray.
//
// This must be called during some calls to ensure that another process does not
// read or write to the SafeArray during editing.
//
// AKA: SafeArrayLock in Windows API.
func Lock(safearray *COMArray) error {
	return com.MaybeError(C.SafeArrayLock(safearray))
}

// Unlock unlocks SafeArray for reading.
//
// AKA: SafeArrayUnlock in Windows API.
func Unlock(safearray *COMArray) error {
	return com.MaybeError(C.SafeArrayUnlock(safearray))
}

// GetPointerOfIndex gets a pointer to an array element.
//
// AKA: SafeArrayPtrOfIndex in Windows API.
func GetPointerOfIndex(safearray *COMArray, index int64) (ref uintptr, err error) {
	err = com.MaybeError(C.SafeArrayPtrOfIndex(safearray, index, unsafe.Pointer(&ref)))
	return
}

// SetInterfaceID sets the GUID of the interface for the specified safe
// array.
//
// AKA: SafeArraySetIID in Windows API.
func SetInterfaceID(safearray *COMArray, interfaceID *com.GUID) error {
	return com.MaybeError(C.SafeArraySetIID(safearray, interfaceID))
}

// ResetDimensions changes the right-most (least significant) bound of the
// specified safe array.
//
// AKA: SafeArrayRedim in Windows API.
func ResetDimensions(safearray *COMArray, bounds *Bounds) error {
	return com.MaybeError(C.SafeArrayRedim(safearray, bounds))
}

// PutElement stores the data element at the specified location in the
// array.
//
// AKA: SafeArrayPutElement in Windows API.
func PutElement(safearray *COMArray, index int64, element interface{}) error {
	return com.MaybeError(C.SafeArrayPutElement(safearray, index, unsafe.Pointer(&element)))
}

// GetRecordInfo accesses IRecordInfo info for custom types.
//
// AKA: SafeArrayGetRecordInfo in Windows API.
//
// XXX: Must implement IRecordInfo interface for this to return.
func GetRecordInfo(safearray *COMArray) (recordInfo interface{}, err error) {
	err = com.MaybeError(C.SafeArrayGetRecordInfo(safearray, unsafe.Pointer(&recordInfo)))
	return
}

// SetRecordInfo mutates IRecordInfo info for custom types.
//
// AKA: SafeArraySetRecordInfo in Windows API.
//
// XXX: Must implement IRecordInfo interface for this to return.
func SetRecordInfo(safearray *COMArray, recordInfo interface{}) error {
	return com.MaybeError(C.SafeArraySetRecordInfo(safearray, unsafe.Pointer(recordInfo)))
}
