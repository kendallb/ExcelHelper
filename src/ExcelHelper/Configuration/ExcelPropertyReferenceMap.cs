/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Reflection;

namespace ExcelHelper.Configuration
{
    /// <summary>
    /// Mapping info for a reference property mapping to a class.
    /// </summary>
    public class ExcelPropertyReferenceMap
    {
        private readonly PropertyInfo _property;

        /// <summary>
        /// Gets the property.
        /// </summary>
        public PropertyInfo Property => _property;

        /// <summary>
        /// Gets the mapping.
        /// </summary>
        public ExcelClassMap Mapping { get; protected set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelPropertyReferenceMap"/> class.
        /// </summary>
        /// <param name="property">The property.</param>
        /// <param name="mapping">The <see cref="ExcelClassMap"/> to use for the reference map.</param>
        public ExcelPropertyReferenceMap(
            PropertyInfo property,
            ExcelClassMap mapping)
        {
            if (mapping == null) {
                throw new ArgumentNullException(nameof(mapping));
            }

            _property = property;
            Mapping = mapping;
        }

        /// <summary>
        /// Get the largest index for the
        /// properties and references.
        /// </summary>
        /// <returns>The max index.</returns>
        internal int GetMaxIndex()
        {
            return Mapping.GetMaxIndex();
        }
    }
}