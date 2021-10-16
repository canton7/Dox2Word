/**
 * @defgroup CommonUnit Common - Unit
 * @ingroup Common
 * 
 * Common Unit is a Unit in the Common module.
 * 
 * @{
 * @file
 */

#include <stdint.h>
#include <stdbool.h>

/**
 * A test structure
 */
typedef struct CommonUnit_Struct_tag
{
	int32_t i; ///< Member 'i'
} CommonUnit_Struct_t;

/**
 * This is the common unit's test file
 * 
 * This is a longer description of the function.
 * 
 * And another paragraph of description.
 * 
 * @param p Some parameters
 * @returns A value
 * 
 * Some more text afterwards.
 */
void CommonUnit_Test(const uint8_t p);

/**
 * Another function in CommonUnit
 *
 * @param in In documentation @c yay
 */
CommonUnit_Struct_t CommonUnit_Another(CommonUnit_Struct_t* in);

/// @}
