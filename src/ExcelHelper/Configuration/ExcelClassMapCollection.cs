/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelHelper.Configuration
{
    /// <summary>
    /// Collection that holds ExcelClassMaps for record types.
    /// </summary>
    public class ExcelClassMapCollection
    {
        private readonly Dictionary<Type, ExcelClassMapBase> _data = new Dictionary<Type, ExcelClassMapBase>();

        /// <summary>
        /// Gets the <see cref="ExcelClassMapBase"/> for the specified record type.
        /// </summary>
        /// <value>
        /// The <see cref="ExcelClassMapBase"/>.
        /// </value>
        /// <param name="type">The record type.</param>
        /// <returns>The <see cref="ExcelClassMapBase"/> for the specified record type.</returns>
        public virtual ExcelClassMapBase this[Type type]
        {
            get
            {
                _data.TryGetValue(type, out var mapBase);
                return mapBase;
            }
        }

        /// <summary>
        /// Adds the specified map for it's record type. If a map
        /// already exists for the record type, the specified
        /// map will replace it.
        /// </summary>
        /// <param name="mapBase">The map.</param>
        internal virtual void Add(
            ExcelClassMapBase mapBase)
        {
            var type = GetGenericExcelClassMapType(mapBase.GetType()).GetGenericArguments().First();

            if (_data.ContainsKey(type)) {
                _data[type] = mapBase;
            } else {
                _data.Add(type, mapBase);
            }
        }

        /// <summary>
        /// Removes the class map.
        /// </summary>
        /// <param name="classMapType">The class map type.</param>
        internal virtual void Remove(
            Type classMapType)
        {
            if (!typeof(ExcelClassMapBase).IsAssignableFrom(classMapType)) {
                throw new ArgumentException("The class map type must inherit from ExcelClassMap.");
            }
            var type = GetGenericExcelClassMapType(classMapType).GetGenericArguments().First();
            _data.Remove(type);
        }

        /// <summary>
        /// Removes all maps.
        /// </summary>
        internal virtual void Clear()
        {
            _data.Clear();
        }

        /// <summary>
        /// Goes up the inheritance tree to find the type instance of ExcelClassMap{}.
        /// </summary>
        /// <param name="type">The type to traverse.</param>
        /// <returns>The type that is ExcelClassMap{}.</returns>
        protected virtual Type GetGenericExcelClassMapType(
            Type type)
        {
            if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(ExcelClassMap<>)) {
                return type;
            }
            return GetGenericExcelClassMapType(type.BaseType);
        }
    }
}
