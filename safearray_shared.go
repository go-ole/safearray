// +build windows,!cgo

package com

import "unsafe"
import "github.com/go-ole/com"

var (
	procSafeArrayAccessData, _        = modoleaut32.FindProc("SafeArrayAccessData")
	procSafeArrayAllocData, _         = modoleaut32.FindProc("SafeArrayAllocData")
	procSafeArrayAllocDescriptor, _   = modoleaut32.FindProc("SafeArrayAllocDescriptor")
	procSafeArrayAllocDescriptorEx, _ = modoleaut32.FindProc("SafeArrayAllocDescriptorEx")
	procSafeArrayCopy, _              = modoleaut32.FindProc("SafeArrayCopy")
	procSafeArrayCopyData, _          = modoleaut32.FindProc("SafeArrayCopyData")
	procSafeArrayCreate, _            = modoleaut32.FindProc("SafeArrayCreate")
	procSafeArrayCreateEx, _          = modoleaut32.FindProc("SafeArrayCreateEx")
	procSafeArrayCreateVector, _      = modoleaut32.FindProc("SafeArrayCreateVector")
	procSafeArrayCreateVectorEx, _    = modoleaut32.FindProc("SafeArrayCreateVectorEx")
	procSafeArrayDestroy, _           = modoleaut32.FindProc("SafeArrayDestroy")
	procSafeArrayDestroyData, _       = modoleaut32.FindProc("SafeArrayDestroyData")
	procSafeArrayDestroyDescriptor, _ = modoleaut32.FindProc("SafeArrayDestroyDescriptor")
	procSafeArrayGetDim, _            = modoleaut32.FindProc("SafeArrayGetDim")
	procSafeArrayGetElement, _        = modoleaut32.FindProc("SafeArrayGetElement")
	procSafeArrayGetElemsize, _       = modoleaut32.FindProc("SafeArrayGetElemsize")
	procSafeArrayGetIID, _            = modoleaut32.FindProc("SafeArrayGetIID")
	procSafeArrayGetLBound, _         = modoleaut32.FindProc("SafeArrayGetLBound")
	procSafeArrayGetUBound, _         = modoleaut32.FindProc("SafeArrayGetUBound")
	procSafeArrayGetVartype, _        = modoleaut32.FindProc("SafeArrayGetVartype")
	procSafeArrayLock, _              = modoleaut32.FindProc("SafeArrayLock")
	procSafeArrayPtrOfIndex, _        = modoleaut32.FindProc("SafeArrayPtrOfIndex")
	procSafeArrayUnaccessData, _      = modoleaut32.FindProc("SafeArrayUnaccessData")
	procSafeArrayUnlock, _            = modoleaut32.FindProc("SafeArrayUnlock")
	procSafeArrayPutElement, _        = modoleaut32.FindProc("SafeArrayPutElement")
	procSafeArrayRedim, _             = modoleaut32.FindProc("SafeArrayRedim")
	procSafeArraySetIID, _            = modoleaut32.FindProc("SafeArraySetIID")
	procSafeArrayGetRecordInfo, _     = modoleaut32.FindProc("SafeArrayGetRecordInfo")
	procSafeArraySetRecordInfo, _     = modoleaut32.FindProc("SafeArraySetRecordInfo")
)

// AccessData returns raw array.
//
// AKA: SafeArrayAccessData in Windows API.
func AccessData(safearray *COMArray) (element uintptr, err error) {
	err = com.HResultToError(procSafeArrayAccessData.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(unsafe.Pointer(&element))))
	return
}

// UnaccessData releases raw array.
//
// AKA: SafeArrayUnaccessData in Windows API.
func UnaccessData(safearray *COMArray) (err error) {
	err = com.HResultToError(procSafeArrayUnaccessData.Call(uintptr(unsafe.Pointer(safearray))))
	return
}

// AllocateArrayData allocates SafeArray.
//
// AKA: SafeArrayAllocData in Windows API.
func AllocateArrayData(safearray *COMArray) (err error) {
	err = com.HResultToError(procSafeArrayAllocData.Call(uintptr(unsafe.Pointer(safearray))))
	return
}

// AllocateArrayDescriptor allocates SafeArray.
//
// AKA: SafeArrayAllocDescriptor in Windows API.
func AllocateArrayDescriptor(dimensions uint32) (safearray *COMArray, err error) {
	err = com.HResultToError(procSafeArrayAllocDescriptor.Call(
		uintptr(dimensions),
		uintptr(unsafe.Pointer(&safearray))))
	return
}

// AllocateArrayDescriptorEx allocates SafeArray.
//
// AKA: SafeArrayAllocDescriptorEx in Windows API.
func AllocateArrayDescriptorEx(variantType com.VariantType, dimensions uint32) (safearray *COMArray, err error) {
	err = com.HResultToError(procSafeArrayAllocDescriptorEx.Call(
		uintptr(variantType),
		uintptr(dimensions),
		uintptr(unsafe.Pointer(&safearray))))
	return
}

// Duplicate returns copy of SafeArray.
//
// AKA: SafeArrayCopy in Windows API.
func Duplicate(original *COMArray) (safearray *COMArray, err error) {
	err = com.HResultToError(procSafeArrayCopy.Call(
		uintptr(unsafe.Pointer(original)),
		uintptr(unsafe.Pointer(&safearray))))
	return
}

// DuplicateData duplicates SafeArray into another SafeArray object.
//
// AKA: SafeArrayCopyData in Windows API.
func DuplicateData(original, duplicate *COMArray) (err error) {
	err = com.HResultToError(procSafeArrayCopyData.Call(
		uintptr(unsafe.Pointer(original)),
		uintptr(unsafe.Pointer(&duplicate))))
	return
}

// CreateArray creates SafeArray.
//
// AKA: SafeArrayCreate in Windows API.
func CreateArray(variantType com.VariantType, dimensions uint32, bounds *Bounds) (safearray *COMArray, err error) {
	sa, _, err := procSafeArrayCreate.Call(
		uintptr(variantType),
		uintptr(dimensions),
		uintptr(unsafe.Pointer(bounds)))
	safearray = (*COMArray)(unsafe.Pointer(&sa))
	return
}

// CreateArrayEx creates SafeArray.
//
// AKA: SafeArrayCreateEx in Windows API.
func CreateArrayEx(variantType com.VariantType, dimensions uint32, bounds *Bounds, extra uintptr) (safearray *COMArray, err error) {
	sa, _, err := procSafeArrayCreateEx.Call(
		uintptr(variantType),
		uintptr(dimensions),
		uintptr(unsafe.Pointer(bounds)),
		extra)
	safearray = (*COMArray)(unsafe.Pointer(sa))
	return
}

// CreateArrayVector creates SafeArray.
//
// AKA: SafeArrayCreateVector in Windows API.
func CreateArrayVector(variantType com.VariantType, lowerBound int32, length uint32) (safearray *COMArray, err error) {
	sa, _, err := procSafeArrayCreateVector.Call(
		uintptr(variantType),
		uintptr(lowerBound),
		uintptr(length))
	safearray = (*COMArray)(unsafe.Pointer(sa))
	return
}

// CreateArrayVectorEx creates SafeArray.
//
// AKA: SafeArrayCreateVectorEx in Windows API.
func CreateArrayVectorEx(variantType com.VariantType, lowerBound int32, length uint32, extra uintptr) (safearray *COMArray, err error) {
	sa, _, err := procSafeArrayCreateVectorEx.Call(
		uintptr(variantType),
		uintptr(lowerBound),
		uintptr(length),
		extra)
	safearray = (*COMArray)(unsafe.Pointer(sa))
	return
}

// Destroy destroys SafeArray object.
//
// AKA: SafeArrayDestroy in Windows API.
func Destroy(safearray *COMArray) error {
	return com.HResultToError(procSafeArrayDestroy.Call(uintptr(unsafe.Pointer(safearray))))
}

// DestroyData destroys SafeArray object.
//
// AKA: SafeArrayDestroyData in Windows API.
func DestroyData(safearray *COMArray) error {
	return com.HResultToError(procSafeArrayDestroyData.Call(uintptr(unsafe.Pointer(safearray))))
}

// DestroyDescriptor destroys SafeArray object.
//
// AKA: SafeArrayDestroyDescriptor in Windows API.
func DestroyDescriptor(safearray *COMArray) error {
	return com.HResultToError(procSafeArrayDestroyDescriptor.Call(uintptr(unsafe.Pointer(safearray))))
}

// GetDimensions is the amount of dimensions in the SafeArray.
//
// SafeArrays may have multiple dimensions. Meaning, it could be
// multidimensional array.
//
// AKA: SafeArrayGetDim in Windows API.
func GetDimensions(safearray *COMArray) (dimensions uint32, err error) {
	l, _, err := procSafeArrayGetDim.Call(uintptr(unsafe.Pointer(safearray)))
	dimensions = *(*uint32)(unsafe.Pointer(l))
	return
}

// GetElementSize is the element size in bytes.
//
// AKA: SafeArrayGetElemsize in Windows API.
func GetElementSize(safearray *COMArray) (length uint32, err error) {
	l, _, err := procSafeArrayGetElemsize.Call(uintptr(unsafe.Pointer(safearray)))
	length = *(*uint32)(unsafe.Pointer(l))
	return
}

// GetElement retrieves element at given index.
func GetElement(safearray *COMArray, index int64) (element interface{}, err error) {
	err = PutElementIn(safearray, index, &element)
	return
}

// PutElementIn retrieves element value at given index.
//
// AKA: SafeArrayGetElement in Windows API.
func PutElementIn(safearray *COMArray, index int64, element interface{}) (err error) {
	err = com.HResultToError(procSafeArrayGetElement.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(index),
		uintptr(unsafe.Pointer(&element))))
	return
}

// GetInterfaceID is the InterfaceID of the elements in the SafeArray.
//
// AKA: SafeArrayGetIID in Windows API.
func GetInterfaceID(safearray *COMArray) (guid *com.GUID, err error) {
	err = com.HResultToError(procSafeArrayGetIID.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(unsafe.Pointer(&guid))))
	return
}

// GetLowerBound returns lower bounds of SafeArray.
//
// SafeArrays may have multiple dimensions. Meaning, it could be
// multidimensional array.
//
// AKA: SafeArrayGetLBound in Windows API.
func GetLowerBound(safearray *COMArray, dimension uint32) (lowerBound int64, err error) {
	err = com.HResultToError(procSafeArrayGetLBound.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(dimension),
		uintptr(unsafe.Pointer(&lowerBound))))
	return
}

// GetUpperBound returns upper bounds of SafeArray.
//
// SafeArrays may have multiple dimensions. Meaning, it could be
// multidimensional array.
//
// AKA: SafeArrayGetUBound in Windows API.
func GetUpperBound(safearray *COMArray, dimension uint32) (upperBound int64, err error) {
	err = com.HResultToError(procSafeArrayGetUBound.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(dimension),
		uintptr(unsafe.Pointer(&upperBound))))
	return
}

// GetVariantType returns data type of SafeArray.
//
// AKA: SafeArrayGetVartype in Windows API.
func GetVariantType(safearray *COMArray) (varType uint16, err error) {
	err = com.HResultToError(procSafeArrayGetVartype.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(unsafe.Pointer(&varType))))
	return
}

// Lock locks SafeArray for reading to modify SafeArray.
//
// This must be called during some calls to ensure that another process does not
// read or write to the SafeArray during editing.
//
// AKA: SafeArrayLock in Windows API.
func Lock(safearray *COMArray) (err error) {
	err = com.HResultToError(procSafeArrayLock.Call(uintptr(unsafe.Pointer(safearray))))
	return
}

// Unlock unlocks SafeArray for reading.
//
// AKA: SafeArrayUnlock in Windows API.
func Unlock(safearray *COMArray) (err error) {
	err = com.HResultToError(procSafeArrayUnlock.Call(uintptr(unsafe.Pointer(safearray))))
	return
}

// GetPointerOfIndex gets a pointer to an array element.
//
// AKA: SafeArrayPtrOfIndex in Windows API.
func GetPointerOfIndex(safearray *COMArray, index int64) (ref uintptr, err error) {
	err = com.HResultToError(procSafeArrayPtrOfIndex.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(index),
		uintptr(unsafe.Pointer(&ref))))
	return
}

// SetInterfaceID sets the GUID of the interface for the specified safe
// array.
//
// AKA: SafeArraySetIID in Windows API.
func SetInterfaceID(safearray *COMArray, interfaceID *com.GUID) (err error) {
	err = com.HResultToError(procSafeArraySetIID.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(unsafe.Pointer(interfaceID))))
	return
}

// ResetDimensions changes the right-most (least significant) bound of the
// specified safe array.
//
// AKA: SafeArrayRedim in Windows API.
func ResetDimensions(safearray *COMArray, bounds *Bounds) error {
	err = com.HResultToError(procSafeArrayRedim.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(unsafe.Pointer(bounds))))
	return
}

// PutElement stores the data element at the specified location in the
// array.
//
// AKA: SafeArrayPutElement in Windows API.
func PutElement(safearray *COMArray, index int64, element interface{}) (err error) {
	err = com.HResultToError(procSafeArrayPutElement.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(index),
		uintptr(unsafe.Pointer(&element))))
	return
}

// GetRecordInfo accesses IRecordInfo info for custom types.
//
// AKA: SafeArrayGetRecordInfo in Windows API.
//
// XXX: Must implement IRecordInfo interface for this to return.
func GetRecordInfo(safearray *COMArray) (recordInfo interface{}, err error) {
	err = com.HResultToError(procSafeArrayGetRecordInfo.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(unsafe.Pointer(&recordInfo))))
	return
}

// SetRecordInfo mutates IRecordInfo info for custom types.
//
// AKA: SafeArraySetRecordInfo in Windows API.
//
// XXX: Must implement IRecordInfo interface for this to return.
func SetRecordInfo(safearray *COMArray, recordInfo interface{}) (err error) {
	err = com.HResultToError(procSafeArraySetRecordInfo.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(unsafe.Pointer(recordInfo))))
	return
}
