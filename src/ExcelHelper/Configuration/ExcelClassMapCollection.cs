/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
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
        private readonly Dictionary<Type, ExcelClassMap> _data = new Dictionary<Type, ExcelClassMap>();

        /// <summary>
        /// Gets the <see cref="ExcelClassMap"/> for the specified record type.
        /// </summary>
        /// <value>
        /// The <see cref="ExcelClassMap"/>.
        /// </value>
        /// <param name="type">The record type.</param>
        /// <returns>The <see cref="ExcelClassMap"/> for the specified record type.</returns>
        public virtual ExcelClassMap this[Type type]
        {
            get
            {
                ExcelClassMap map;
                _data.TryGetValue(type, out map);
                return map;
            }
        }

        /// <summary>
        /// Adds the specified map for it's record type. If a map
        /// already exists for the record type, the specified
        /// map will replace it.
        /// </summary>
        /// <param name="map">The map.</param>
        internal virtual void Add(
            ExcelClassMap map)
        {
            var type = GetGenericExcelClassMapType(map.GetType()).GetGenericArguments().First();

            if (_data.ContainsKey(type)) {
                _data[type] = map;
            } else {
                _data.Add(type, map);
            }
        }

        /// <summary>
        /// Removes the class map.
        /// </summary>
        /// <param name="classMapType">The class map type.</param>
        internal virtual void Remove(
            Type classMapType)
        {
            if (!typeof(ExcelClassMap).IsAssignableFrom(classMapType)) {
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
