package safearray

import (
	"errors"
	"reflect"
	"unsafe"

	"github.com/go-ole/com"
	"github.com/go-ole/idispatch"
	"github.com/go-ole/iunknown"
)

var (
	NotSliceError              = errors.New("must be a slice")
	VariantNotImplementedError = errors.New("variant type is not implemented")
	VariantNotSupportedError   = errors.New("variant type is not supported")
	SafeArrayVectorFailedError = errors.New("could not convert slice to SafeArray")
)

// COMArray is how COM handles arrays.
type COMArray struct {
	Dimensions   uint16
	FeaturesFlag uint16
	ElementsSize uint32
	LocksAmount  uint32
	Data         uint32

	// This must hold the bytes for two Bounds objects. Use binary.Read() to
	// get the contents.
	Bounds [16]byte
}

// Bounds defines the array boundaries.
type Bounds struct {
	Elements   uint32
	LowerBound int32
}

// Array is wrapper for COMArray with helpers.
//
// It is recommended that you use this type instead of the COMArray, because
// the bounds is a pointer to the SafeArrayBounds and not referenced directly.
type Array struct {
	Array *COMArray

	// bounds contains a mirror of the COMArray.Bounds in the Go type.
	bounds [2]Bounds
}

// Destroy SafeArray object.
func (sa *Array) Destroy() error {
	return Destroy(sa.Array)
}

// DestroyData removes safe array data.
func (sa *Array) DestroyData() error {
	return DestroyData(sa.Array)
}

// DestroyDescriptor removes safe array descriptor.
func (sa *Array) DestroyDescriptor() error {
	return DestroyDescriptor(sa.Array)
}

// Duplicate SafeArray to another SafeArray.
//
// This copies the underlying COMArray object into another Array object.
func (sa *Array) Duplicate() (*Array, error) {
	saCopy, err := Duplicate(sa.Array)
	return &Array{saCopy}, err
}

// DuplicateDataTo takes current SafeArray Data and copies to given SafeArray.
//
// This copies the underlying COMArray data into another SafeArray
// COMArray object.
func (sa *Array) DuplicateDataTo(duplicate *Array) error {
	return DuplicateData(sa.Array, duplicate.Array)
}

// Dimensions is the total number of array of arrays.
//
// For example is dimensions returns 3, then you have:
//
//     array[0][]
//     array[1][]
//     array[2][]
//
// And so on for other lengths.
func (sa *Array) Dimensions() (uint32, error) {
	return GetDimensions(sa.Array)
}

// ResetDimensions resets the bounds of the SafeArray.
//
// If the bounds is less than the current, then memory will be automatically
// freed. If the bounds is more than the current, then memory will be
// automatically allocated.
func (sa *Array) ResetDimensions(bounds []Bounds) error {
	sa.bounds = bounds
	return ResetDimensions(sa.Array, &sa.bounds[0])
}

// ElementSize is the type's size.
func (sa *Array) ElementSize() (uint32, error) {
	return GetElementSize(sa.Array)
}

// Length returns total elements for SafeArray.
func (sa *Array) Length() (totalElements int64, err error) {
	totalElements = 0
	dimensions, err := sa.Dimensions()
	if err != nil {
		return
	}

	for dimension := uint32(1); dimension <= dimensions; dimension++ {
		length, err := sa.DimensionLength(dimension)
		if err != nil {
			return
		}
		totalElements += length
	}

	return
}

// DimensionLength returns total elements for given dimension.
//
// Dimensions start at 1, this will only be corrected if you enter '0'.
func (sa *Array) DimensionLength(dimension uint32) (totalElements int64, err error) {
	if dimension < 1 {
		dimension = 1
	}

	// Get array bounds
	var LowerBounds int64
	var UpperBounds int64

	LowerBounds, err = GetLowerBound(sa.Array, dimension)
	if err != nil {
		return
	}

	UpperBounds, err = GetUpperBound(sa.Array, dimension)
	if err != nil {
		return
	}

	totalElements = UpperBounds - LowerBounds + 1
	return
}

// SetElementAt with element value at index.
//
// XXX: Index must be defined on how it works with multidimensional arrays.
func (sa *Array) SetElementAt(index int64, element interface{}) error {
	return PutElement(sa.Array, index, &element)
}

// ElementAt returns element at index.
//
// Returned value will need to be converted to the type you require, because it
// is an interface{}.
//
// XXX: Index must be defined on how it works with multidimensional arrays.
func (sa *Array) ElementAt(index int64) (interface{}, error) {
	return GetElement(sa.Array, index)
}

// ElementFor puts element value into given element.
//
// You do not need to convert element. It will be typed to the interface. This
// is an unsafe operation. Element must be passed by reference.
//
// XXX: Index must be defined on how it works with multidimensional arrays.
func (sa *Array) ElementFor(index int64, element interface{}) error {
	return PutElementIn(sa.Array, index, &element)
}

// SetInterfaceID sets the IID for the COM array.
//
// This is only used when serving COM arrays to clients.
func (sa *Array) SetInterfaceID(interfaceID *com.GUID) error {
	return SetInterfaceID(sa.Array, &interfaceID)
}

// InterfaceID may return the IID, if the array type is a COM object.
func (sa *Array) InterfaceID() (*com.GUID, error) {
	return GetInterfaceID(sa.Array)
}

// VariantType returns the variant type, if there is one available.
//
// Flag com.HasVariantSafeArrayMask must be set.
func (sa *Array) VariantType() (varType com.VariantType, err error) {
	vt, err := GetVariantType(sa.Array)
	varType = com.VariantType(vt)
	return
}

// Lock SafeArray for modification.
func (sa *Array) Lock() error {
	return Lock(sa.Array)
}

// Unlock SafeArray for reading.
func (sa *Array) Unlock() error {
	return Unlock(sa.Array)
}

// RecordInfo retrieves IRecordInfo for SafeArray.
//
// XXX: Must implement IRecordInfo interface for this to return.
func (sa *Array) RecordInfo() (interface{}, error) {
	return GetRecordInfo(sa.Array)
}

// SetRecordInfo sets IRecordInfo for SafeArray.
//
// XXX: Must implement IRecordInfo interface for this to return.
func (sa *Array) SetRecordInfo(info interface{}) error {
	return SetRecordInfo(sa.Array, info)
}

// PutInArray converts SafeArray data in to arbitrary type slice.
//
// This works on both single dimensional and multidimensional arrays. It will
// convert multidimensional to single dimensional arrays. This will not change
// in the future. A separate method exists for returning a multidimensional
// array.
func (sa *Array) PutInArray(slice interface{}) (err error) {
	if !IsSlice(slice) {
		err = NotSliceError
		return
	}

	dimensions, err := GetDimensions(sa.Array)
	if err != nil {
		return
	}

	length, err := sa.Length()
	if err != nil {
		return
	}

	kind := reflect.ValueOf(slice).Kind()

	if dimensions == 1 && kind != reflect.String {
		err = MarshalArray(sa.Array, length, &slice)
		return
	}

	t := reflect.TypeOf(slice)

	for i := int64(0); i < length; i++ {
		if kind != string {
			element := reflect.New(t).Interface()
			err = PutElementIn(sa.Array, i, &element)
			if err != nil {
				return
			}
			*slice = append(slice, element)
		} else {
			element, err := GetElementString(sa.Array, i)
			if err != nil {
				return
			}
			*slice = append(slice, element)
		}
	}
}

// ToArray returns array slice based on the supported variant type.
//
// If there is no variant type, then an error will be returned.
//
// You must also convert the returned value to whatever slice type it should be.
//
//     raw, err := array.ToArray()
//     slice, ok := raw.([]byte)
//
// This must be done because the returned type is a slice of interface{}.
func (sa *Array) ToArray() (slice interface{}, err error) {
	vt, err := sa.VariantType()
	if err != nil {
		return
	}

	// Must not have VT_ARRAY and VT_BYREF flags set.
	// Must not be VT_EMPTY and VT_NULL.

	switch vt {
	case com.Float32VariantType:
		slice = make([]float32, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.Float64VariantType:
		slice = make([]float64, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.CurrencyVariantType:
		slice = make([]*com.Currency, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.DateVariantType:
		err = VariantNotImplementedError
	case com.BinaryStringVariantType, com.ClassIDVariantType:
		slice = make([]string, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.IDispatchVariantType:
		slice = make([]*idispatch.Dispatch, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.ErrorVariantType:
		err = VariantNotImplementedError
	case com.BoolVariantType:
		slice = make([]uint16, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.VariantVariantType:
		slice = make([]*com.Variant, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.IUnknownVariantType:
		slice = make([]*iunknown.Unknown, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.DecimalVariantType:
		slice = make([]*com.Decimal, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.Integer8VariantType:
		slice = make([]int8, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.UInteger8VariantType:
		slice = make([]uint8, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.Integer16VariantType:
		slice = make([]int16, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.UInteger16VariantType:
		slice = make([]uint16, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.Integer32VariantType:
		slice = make([]int32, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.UInteger32VariantType:
		slice = make([]uint32, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.Integer64VariantType:
		slice = make([]int64, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.UInteger64VariantType:
		slice = make([]uint64, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.IntegerVariantType:
		// Warning: This must match the architecture of the application you wish
		// to access.
		slice = make([]int, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.UIntegerVariantType:
		// Warning: This must match the architecture of the application you wish
		// to access.
		slice = make([]uint, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.HResultVariantType:
		// Warning: This must match the architecture of the application you wish
		// to access.
		slice = make([]uintptr, sa.Length())
		err = sa.PutInArray(&slice)
		// TODO: Need to turn HResult into OleError.
		return
	case com.PointerVariantType:
		slice = make([]unsafe.Pointer, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.SafeArrayVariantType:
		slice = make([]*COMArray, sa.Length())
		err = sa.PutInArray(&slice)
		// Need to turn into Array objects
		return
	case com.CArrayVariantType:
		// TODO: Complete
		err = VariantNotImplementedError
	case com.ANSIStringVariantType:
		// TODO: Complete
		err = VariantNotImplementedError
	case com.UnicodeStringVariantType:
		// TODO: Complete
		err = VariantNotImplementedError
	case com.RecordVariantType:
		// TODO: Complete
		err = VariantNotImplementedError
	case com.IntegerPointerVariantType, com.UIntegerPointerVariantType:
		slice = make([]uintptr, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.FileTimeVariantType:
		slice = make([]*com.FileTime, sa.Length())
		err = sa.PutInArray(&slice)
		return
	case com.ClipboardFormatVariantType:
		// TODO: Complete
		err = VariantNotImplementedError
	default:
		err = VariantNotSupportedError
	}

	return
}
